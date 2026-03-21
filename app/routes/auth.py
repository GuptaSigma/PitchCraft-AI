from flask import Blueprint, request, jsonify
from functools import wraps
import bcrypt
import jwt
import datetime
import os
import random
import traceback
import importlib
import requests as http_requests

auth_bp = Blueprint('auth', __name__)

print("✅ auth_bp Blueprint created")

JWT_SECRET = os.getenv('JWT_SECRET_KEY')
if not JWT_SECRET:
    raise RuntimeError("❌ JWT_SECRET_KEY missing from .env file!")

JWT_ALGORITHM = 'HS256'
JWT_EXPIRATION_HOURS = 24
OTP_EXPIRATION_MINUTES = int(os.getenv('OTP_EXPIRATION_MINUTES', '10'))
GOOGLE_CLIENT_ID = os.getenv('GOOGLE_CLIENT_ID', '').strip()
GOOGLE_APPS_SCRIPT_WEBHOOK_URL = os.getenv(
    'GOOGLE_APPS_SCRIPT_WEBHOOK_URL',
    'https://script.google.com/macros/s/AKfycbxZdTmijo5OZhZcG2W7XFLQCmUQtwYDwfLGKri43qeeHpoeSC-JN1d934qCp6kT1Sj1/exec'
).strip()


def execute_query(query, params=None, fetch=False):
    database_module = importlib.import_module('app.models.database')
    return database_module.execute_query(query, params, fetch)


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')


def verify_password(password: str, hashed: str) -> bool:
    return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))


def generate_token(user_id: int, email: str, name: str) -> str:
    payload = {
        'user_id': user_id,
        'email': email,
        'name': name,
        'exp': datetime.datetime.utcnow() + datetime.timedelta(hours=JWT_EXPIRATION_HOURS),
        'iat': datetime.datetime.utcnow()
    }
    return jwt.encode(payload, JWT_SECRET, algorithm=JWT_ALGORITHM)


def verify_token(token: str):
    try:
        return jwt.decode(token, JWT_SECRET, algorithms=[JWT_ALGORITHM])
    except jwt.ExpiredSignatureError:
        print('❌ Token expired')
        return None
    except jwt.InvalidTokenError:
        print('❌ Invalid token')
        return None


def generate_otp() -> str:
    return f"{random.randint(0, 999999):06d}"


def send_otp_email(email: str, name: str, otp_code: str) -> tuple[bool, str]:
    if not GOOGLE_APPS_SCRIPT_WEBHOOK_URL:
        return False, 'Google Apps Script webhook URL missing'

    payload = {
        'to': email,
        'name': name,
        'otp': otp_code,
        'subject': 'PitchCraft AI OTP Verification',
        'message': (
            f'Hello {name or "User"},\\n\\n'
            f'Your PitchCraft AI OTP is: {otp_code}\\n'
            f'This OTP will expire in {OTP_EXPIRATION_MINUTES} minutes.\\n\\n'
            'If you did not request this, please ignore this email.'
        )
    }

    try:
        response = http_requests.post(
            GOOGLE_APPS_SCRIPT_WEBHOOK_URL,
            json=payload,
            timeout=20,
            headers={'Content-Type': 'application/json'}
        )
        if response.status_code == 200:
            return True, 'OTP sent successfully'
        return False, f'Apps Script error: {response.status_code}'
    except Exception as exc:
        print(f'❌ OTP send error: {exc}')
        return False, 'Failed to send OTP email'


@auth_bp.route('/config', methods=['GET'])
def auth_config():
    return jsonify({
        'success': True,
        'googleClientId': GOOGLE_CLIENT_ID,
        'otpEnabled': bool(GOOGLE_APPS_SCRIPT_WEBHOOK_URL)
    }), 200


def token_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        token = None
        if 'Authorization' in request.headers:
            auth_header = request.headers['Authorization']
            if auth_header.startswith('Bearer '):
                token = auth_header.split(' ')[1]

        if not token and 'token' in request.args:
            token = request.args.get('token')

        if not token:
            return jsonify({'success': False, 'error': 'Token is missing'}), 401

        payload = verify_token(token)
        if not payload:
            return jsonify({'success': False, 'error': 'Invalid or expired token'}), 401

        return f(payload['user_id'], *args, **kwargs)

    return decorated


def validate_signup_payload(data):
    if not data or not data.get('email') or not data.get('password') or not data.get('name'):
        return 'Name, email, and password required'

    email = data['email'].lower().strip()
    password = data['password']
    name = data['name'].strip()

    if '@' not in email or '.' not in email:
        return 'Invalid email format'
    if len(password) < 6:
        return 'Password must be at least 6 characters'
    if not name:
        return 'Name is required'
    return None


@auth_bp.route('/signup', methods=['POST', 'OPTIONS'])
def signup():
    if request.method == 'OPTIONS':
        return '', 204

    try:
        data = request.get_json() or {}
        otp_code = (data.get('otp') or '').strip()

        if otp_code:
            email = (data.get('email') or '').lower().strip()
            if not email:
                return jsonify({'success': False, 'error': 'Email is required for OTP verification'}), 400

            rows = execute_query(
                """
                SELECT * FROM otp_verifications
                WHERE email = %s AND purpose = 'signup' AND consumed_at IS NULL
                ORDER BY id DESC
                """,
                (email,),
                fetch=True
            )

            if not rows:
                return jsonify({'success': False, 'error': 'OTP not found. Please request a new OTP.'}), 404

            otp_record = rows[0]
            if otp_record['expires_at'] < datetime.datetime.utcnow():
                return jsonify({'success': False, 'error': 'OTP expired. Please request a new OTP.'}), 400

            if otp_record['otp_code'] != otp_code:
                return jsonify({'success': False, 'error': 'Invalid OTP'}), 400

            existing_user = execute_query(
                'SELECT id FROM users WHERE email = %s',
                (email,),
                fetch=True
            )
            if existing_user:
                return jsonify({'success': False, 'error': 'Email already registered'}), 409

            user_id = execute_query(
                """
                INSERT INTO users (name, email, password, auth_provider, email_verified)
                VALUES (%s, %s, %s, %s, %s)
                """,
                (
                    otp_record['name'],
                    email,
                    otp_record['password_hash'],
                    'password',
                    True
                )
            )

            execute_query(
                'UPDATE otp_verifications SET consumed_at = %s WHERE id = %s',
                (datetime.datetime.utcnow(), otp_record['id'])
            )

            token = generate_token(user_id, email, otp_record['name'])
            return jsonify({
                'success': True,
                'message': 'Account created successfully',
                'token': token,
                'user': {'id': user_id, 'name': otp_record['name'], 'email': email}
            }), 201

        validation_error = validate_signup_payload(data)
        if validation_error:
            return jsonify({'success': False, 'error': validation_error}), 400

        email = data['email'].lower().strip()
        password = data['password']
        name = data['name'].strip()

        existing_user = execute_query(
            'SELECT id FROM users WHERE email = %s',
            (email,),
            fetch=True
        )
        if existing_user:
            return jsonify({'success': False, 'error': 'Email already registered'}), 409

        otp_code = generate_otp()
        password_hash = hash_password(password)
        expires_at = datetime.datetime.utcnow() + datetime.timedelta(minutes=OTP_EXPIRATION_MINUTES)

        execute_query(
            "DELETE FROM otp_verifications WHERE email = %s AND purpose = 'signup' AND consumed_at IS NULL",
            (email,)
        )
        execute_query(
            """
            INSERT INTO otp_verifications (email, name, password_hash, otp_code, purpose, expires_at)
            VALUES (%s, %s, %s, %s, 'signup', %s)
            """,
            (email, name, password_hash, otp_code, expires_at)
        )

        sent, message = send_otp_email(email, name, otp_code)
        if not sent:
            return jsonify({'success': False, 'error': message}), 500

        return jsonify({
            'success': True,
            'requiresOtp': True,
            'message': f'OTP sent to {email}',
            'expiresInMinutes': OTP_EXPIRATION_MINUTES
        }), 200

    except Exception as exc:
        print(f'❌ Signup DB error: {exc}')
        return jsonify({'success': False, 'error': 'Database error'}), 500
    except Exception as exc:
        print(f'❌ Signup error: {exc}')
        traceback.print_exc()
        return jsonify({'success': False, 'error': 'Signup failed'}), 500


@auth_bp.route('/login', methods=['POST', 'OPTIONS'])
def login():
    if request.method == 'OPTIONS':
        return '', 204

    try:
        data = request.get_json() or {}
        email = (data.get('email') or '').lower().strip()
        password = data.get('password') or ''

        if not email or not password:
            return jsonify({'success': False, 'error': 'Email and password required'}), 400

        users = execute_query(
            'SELECT id, name, email, password, auth_provider FROM users WHERE email = %s',
            (email,),
            fetch=True
        )

        if not users:
            return jsonify({'success': False, 'error': 'Invalid email or password'}), 401

        user = users[0]
        if not user['password']:
            return jsonify({'success': False, 'error': 'Use Google sign in for this account'}), 400

        if not verify_password(password, user['password']):
            return jsonify({'success': False, 'error': 'Invalid email or password'}), 401

        token = generate_token(user['id'], user['email'], user['name'])
        return jsonify({
            'success': True,
            'message': 'Login successful',
            'token': token,
            'user': {'id': user['id'], 'name': user['name'], 'email': user['email']}
        }), 200

    except Exception as exc:
        print(f'❌ Login DB error: {exc}')
        return jsonify({'success': False, 'error': 'Database error'}), 500
    except Exception as exc:
        print(f'❌ Login error: {exc}')
        traceback.print_exc()
        return jsonify({'success': False, 'error': 'Login failed'}), 500


@auth_bp.route('/google-login', methods=['POST', 'OPTIONS'])
def google_login():
    if request.method == 'OPTIONS':
        return '', 204

    if not GOOGLE_CLIENT_ID:
        return jsonify({'success': False, 'error': 'Google Client ID not configured'}), 500

    try:
        id_token = importlib.import_module('google.oauth2.id_token')
        google_requests = importlib.import_module('google.auth.transport.requests')

        data = request.get_json() or {}
        credential = data.get('credential')
        if not credential:
            return jsonify({'success': False, 'error': 'Google credential is required'}), 400

        idinfo = id_token.verify_oauth2_token(
            credential,
            google_requests.Request(),
            GOOGLE_CLIENT_ID
        )

        email = (idinfo.get('email') or '').lower().strip()
        google_id = idinfo.get('sub')
        name = idinfo.get('name') or email.split('@')[0]
        email_verified = bool(idinfo.get('email_verified'))

        if not email or not google_id:
            return jsonify({'success': False, 'error': 'Invalid Google account data'}), 400

        users = execute_query(
            'SELECT id, name, email, google_id FROM users WHERE email = %s',
            (email,),
            fetch=True
        )

        if users:
            user = users[0]
            execute_query(
                """
                UPDATE users
                SET google_id = %s, auth_provider = %s, email_verified = %s, name = %s
                WHERE id = %s
                """,
                (google_id, 'google', email_verified, name, user['id'])
            )
            user_id = user['id']
        else:
            user_id = execute_query(
                """
                INSERT INTO users (name, email, password, google_id, auth_provider, email_verified)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                (name, email, None, google_id, 'google', email_verified)
            )

        token = generate_token(user_id, email, name)
        return jsonify({
            'success': True,
            'message': 'Google login successful',
            'token': token,
            'user': {'id': user_id, 'name': name, 'email': email}
        }), 200

    except ValueError as exc:
        print(f'❌ Google token verification error: {exc}')
        return jsonify({'success': False, 'error': 'Invalid Google token'}), 401
    except ImportError:
        return jsonify({'success': False, 'error': 'google-auth package not installed'}), 500
    except Exception as exc:
        print(f'❌ Google login DB error: {exc}')
        return jsonify({'success': False, 'error': 'Database error'}), 500
    except Exception as exc:
        print(f'❌ Google login error: {exc}')
        traceback.print_exc()
        return jsonify({'success': False, 'error': 'Google login failed'}), 500


@auth_bp.route('/verify', methods=['GET', 'OPTIONS'])
def verify_route():
    if request.method == 'OPTIONS':
        return '', 204

    try:
        auth_header = request.headers.get('Authorization')
        if not auth_header or not auth_header.startswith('Bearer '):
            return jsonify({'success': False, 'error': 'No token provided'}), 401

        token = auth_header.split(' ')[1]
        payload = verify_token(token)
        if not payload:
            return jsonify({'success': False, 'error': 'Invalid or expired token'}), 401

        return jsonify({
            'success': True,
            'user': {
                'id': payload['user_id'],
                'email': payload['email'],
                'name': payload['name']
            }
        }), 200

    except Exception as exc:
        print(f'❌ Verify error: {exc}')
        return jsonify({'success': False, 'error': 'Token verification failed'}), 500


@auth_bp.route('/logout', methods=['POST', 'OPTIONS'])
def logout():
    if request.method == 'OPTIONS':
        return '', 204
    return jsonify({'success': True, 'message': 'Logged out successfully'}), 200


print('✅ Auth routes: /config, /signup, /login, /google-login, /verify, /logout')
