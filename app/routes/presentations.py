from flask import Blueprint, request, jsonify, send_file
import json
import datetime
import io
import traceback
import importlib

from app.services.pptx_service import PPTXService
from app.routes.auth import token_required

# PDF Service Import (Optional)
try:
    from app.services.pdf_services import PDFService
    print("✅ PDF Service imported")
except Exception as e:
    PDFService = None
    print(f"⚠️ PDF Service unavailable: {e}")

# DOC Service Import (Optional)
try:
    from app.services.doc_services import DOCService
    print("✅ DOC Service imported")
except Exception as e:
    DOCService = None
    print(f"⚠️ DOC Service unavailable: {e}")

# DOCX Service Import (Optional)
try:
    from app.services.docx_services import DOCXService
    print("✅ DOCX Service imported")
except Exception as e:
    DOCXService = None
    print(f"⚠️ DOCX Service unavailable: {e}")

# AI Service Import (Safe)
try: 
    from app.services.ai_service import ai_service
    print("✅ AI Service imported in presentations routes")
except Exception as e: 
    ai_service = None
    print(f"⚠️ AI Service unavailable: {e}")

# Blueprint Registration
presentations_bp = Blueprint('presentations', __name__)
print("✅ presentations_bp Blueprint created")

DAILY_TOKEN_LIMIT = 400
TOKENS_PER_SLIDE = 5


def execute_query(query, params=None, fetch=False):
    database_module = importlib.import_module('app.models.database')
    return database_module.execute_query(query, params, fetch)


def get_connection():
    database_module = importlib.import_module('app.models.database')
    return database_module.get_connection()


def get_daily_token_usage(user_id):
    usage_date = datetime.date.today()
    rows = execute_query(
        "SELECT tokens_used FROM user_daily_usage WHERE user_id = %s AND usage_date = %s",
        (user_id, usage_date),
        fetch=True
    ) or []
    tokens_used = int(rows[0]['tokens_used']) if rows else 0
    tokens_remaining = max(0, DAILY_TOKEN_LIMIT - tokens_used)
    return {
        'usage_date': usage_date.isoformat(),
        'daily_limit': DAILY_TOKEN_LIMIT,
        'tokens_used': tokens_used,
        'tokens_remaining': tokens_remaining,
        'tokens_per_slide': TOKENS_PER_SLIDE
    }


def reserve_daily_tokens(user_id, tokens_needed):
    usage_date = datetime.date.today()
    conn = None
    cursor = None
    try:
        conn = get_connection()
        cursor = conn.cursor(dictionary=True)
        cursor.execute(
            "SELECT tokens_used FROM user_daily_usage WHERE user_id = %s AND usage_date = %s FOR UPDATE",
            (user_id, usage_date)
        )
        row = cursor.fetchone()
        current_used = int(row['tokens_used']) if row else 0
        updated_used = current_used + tokens_needed

        if updated_used > DAILY_TOKEN_LIMIT:
            conn.rollback()
            return {
                'allowed': False,
                'daily_limit': DAILY_TOKEN_LIMIT,
                'tokens_used': current_used,
                'tokens_remaining': max(0, DAILY_TOKEN_LIMIT - current_used),
                'tokens_needed': tokens_needed,
                'retry_message': 'Token limit finished for today. Tomorrow we come back.'
            }

        if row:
            cursor.execute(
                "UPDATE user_daily_usage SET tokens_used = %s WHERE user_id = %s AND usage_date = %s",
                (updated_used, user_id, usage_date)
            )
        else:
            cursor.execute(
                "INSERT INTO user_daily_usage (user_id, usage_date, tokens_used) VALUES (%s, %s, %s)",
                (user_id, usage_date, tokens_needed)
            )

        conn.commit()
        return {
            'allowed': True,
            'daily_limit': DAILY_TOKEN_LIMIT,
            'tokens_used': updated_used,
            'tokens_remaining': max(0, DAILY_TOKEN_LIMIT - updated_used),
            'tokens_needed': tokens_needed
        }
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


@presentations_bp.route('/quota', methods=['GET'])
@token_required
def get_quota(user_id):
    try:
        quota = get_daily_token_usage(user_id)
        return jsonify({'success': True, 'quota': quota}), 200
    except Exception as e:
        print(f"❌ quota error: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

# GET /api/presentations/  → List All Presentations (Dashboard)
@presentations_bp.route('/', methods=['GET'])
@token_required
def list_presentations(user_id):
    """
    Fetch all presentations for dashboard grid view
    ✅ UPDATED: Filter by user_id
    """
    try: 
        rows = execute_query(
            """
            SELECT id, title, prompt, slides_count, theme, style, created_at, updated_at
            FROM presentations
            WHERE user_id = %s
            ORDER BY created_at DESC
            """,
            (user_id,),
            fetch=True
        ) or []

        for r in rows:
            if isinstance(r.get('created_at'), (datetime.datetime, datetime.date)):
                r['created_at'] = r['created_at'].isoformat()
            if isinstance(r.get('updated_at'), (datetime.datetime, datetime.date)):
                r['updated_at'] = r['updated_at'].isoformat() if r['updated_at'] else None

        return jsonify({'success': True, 'presentations': rows}), 200

    except Exception as e:
        print(f"❌ list_presentations error: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

# GET /api/presentations/<id>  → View Single Presentation (FULL SLIDES)
@presentations_bp.route('/<int:pres_id>', methods=['GET'])
@token_required
def get_presentation(user_id, pres_id):
    """
    Fetch single presentation with FULL content (all slides).
    ✅ UPDATED: Verify ownership
    """
    try:
        print(f"📥 Fetching presentation ID: {pres_id} for User: {user_id}")

        rows = execute_query(
            """
            SELECT id, title, prompt, slides_count, content, theme, style, 
                   language, created_at, updated_at 
            FROM presentations 
            WHERE id = %s AND user_id = %s
            """,
            (pres_id, user_id),
            fetch=True
        )

        if not rows:
            return jsonify({'success': False, 'error': 'Presentation not found'}), 404

        pres = rows[0]

        # Parse JSON 'content' into slides array
        slides = []
        try:
            raw = pres.get('content')
            if raw is None:
                slides = []
            elif isinstance(raw, str):
                slides = json.loads(raw) if raw.strip() else []
            else: 
                slides = raw or []
        except Exception as e:
            print(f"⚠️ JSON parse error for pres_id={pres_id}: {e}")
            slides = []

        # Build response object
        response = dict(pres)
        response.pop('content', None)
        response['slides'] = slides

        # Format datetimes
        if isinstance(response.get('created_at'), (datetime.datetime, datetime.date)):
            response['created_at'] = response['created_at'].isoformat()
        if isinstance(response.get('updated_at'), (datetime.datetime, datetime.date)):
            response['updated_at'] = response['updated_at'].isoformat() if response['updated_at'] else None

        print(f"✅ Found {len(slides)} slides for '{response.get('title')}' with theme: {response.get('theme', 'dialogue')}")
        return jsonify({'success': True, 'presentation': response}), 200

    except Exception as e:
        print(f"❌ get_presentation error: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

# PUT /api/presentations/<id>  → UPDATE PRESENTATION (SAVE CHANGES)
@presentations_bp.route('/<int:pres_id>', methods=['PUT'])
@token_required
def update_presentation(user_id, pres_id):
    """
    Save manual edits made by the user on the frontend.
    Updates the 'content' JSON in the database.
    ✅ UPDATED: Verify ownership
    """
    try:
        data = request.get_json() or {}
        slides = data.get('slides')
        theme = data.get('theme')
        title = data.get('title')

        if slides is None or not isinstance(slides, list):
            return jsonify({'success': False, 'error': 'Invalid slides data'}), 400

        if not isinstance(theme, str) or not theme.strip():
            return jsonify({'success': False, 'error': 'Theme is required'}), 400

        content_json = json.dumps(slides, ensure_ascii=False)
        normalized_theme = theme.lower().strip()

        valid_themes = ['dialogue', 'alien', 'wine', 'snowball', 'petrol', 'piano', 'sunset', 'midnight']
        if normalized_theme not in valid_themes:
            return jsonify({'success': False, 'error': f'Invalid theme. Valid options: {", ".join(valid_themes)}'}), 400

        conn = None
        cursor = None

        try:
            conn = get_connection()
            cursor = conn.cursor(dictionary=True)

            cursor.execute("SELECT theme, title FROM presentations WHERE id = %s AND user_id = %s", (pres_id, user_id))
            before_row = cursor.fetchone() or {}
            
            if not before_row:
                return jsonify({'success': False, 'error': 'Access denied or presentation not found'}), 403

            print(f"🧪 [SAVE] Before update (id={pres_id}) theme={before_row.get('theme')} title={before_row.get('title')}")

            cursor.execute(
                """
                UPDATE presentations
                SET content = %s,
                    theme = %s,
                    title = COALESCE(%s, title),
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = %s AND user_id = %s
                """,
                (content_json, normalized_theme, title, pres_id, user_id)
            )

            if cursor.rowcount == 0:
                conn.rollback()
                return jsonify({'success': False, 'error': 'Presentation not found'}), 404

            conn.commit()

            cursor.execute("SELECT theme, title FROM presentations WHERE id = %s AND user_id = %s", (pres_id, user_id))
            after_row = cursor.fetchone() or {}
            print(f"✅ [SAVE] After update (id={pres_id}) theme={after_row.get('theme')} title={after_row.get('title')}")
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

        return jsonify({'success': True, 'message': 'Presentation updated successfully'}), 200

    except Exception as e: 
        print(f"❌ Update error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

# PUT /api/presentations/<id>/theme  → UPDATE THEME ONLY
@presentations_bp.route('/<int:pres_id>/theme', methods=['PUT'])
@token_required
def update_theme(user_id, pres_id):
    """
    Update presentation theme only (for theme dropdown change)
    ✅ NEW ROUTE: Dedicated theme update endpoint
    ✅ UPDATED: Auth integrated
    """
    try:
        data = request.json or {}
        theme = data.get('theme')

        if not theme:
            return jsonify({'success': False, 'error': 'Theme is required'}), 400

        # Validate theme
        valid_themes = ['dialogue', 'alien', 'wine', 'snowball', 'petrol', 'piano', 'sunset', 'midnight']
        normalized_theme = str(theme).lower().strip()
        if normalized_theme not in valid_themes:
            return jsonify({'success': False, 'error': f'Invalid theme. Valid options: {", ".join(valid_themes)}'}), 400

        conn = None
        cursor = None

        try:
            conn = get_connection()
            cursor = conn.cursor(dictionary=True)

            # Debug: confirm current database
            cursor.execute("SELECT current_database() AS db")
            db_row = cursor.fetchone() or {}
            print(f"🧭 [THEME] Connected DB: {db_row.get('db')}")

            # Debug: theme before update
            cursor.execute(
                "SELECT theme FROM presentations WHERE id = %s AND user_id = %s",
                (pres_id, user_id)
            )
            before_row = cursor.fetchone() or {}
            
            if not before_row:
                 return jsonify({'success': False, 'error': 'Access denied or presentation not found'}), 403

            print(f"🧪 [THEME] Before update (id={pres_id}): {before_row.get('theme')}")

            # Update theme in database
            cursor.execute(
                "UPDATE presentations SET theme = %s WHERE id = %s AND user_id = %s",
                (normalized_theme, pres_id, user_id)
            )
            print(f"🧾 [THEME] Rowcount after update: {cursor.rowcount}")

            if cursor.rowcount == 0:
                conn.rollback()
                return jsonify({'success': False, 'error': 'Presentation not found'}), 404

            # Ensure persistence
            conn.commit()

            # Re-query to confirm persistence
            cursor.execute(
                "SELECT id, theme FROM presentations WHERE id = %s AND user_id = %s",
                (pres_id, user_id)
            )
            row = cursor.fetchone()

            saved_theme = (row or {}).get('theme')
            print(f"🎨 [THEME] After update (id={pres_id}): {saved_theme}")

        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

        return jsonify({
            'success': True,
            'message': 'Theme updated successfully',
            'theme': saved_theme
        }), 200

    except Exception as e:
        print(f"❌ Theme update error: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

# POST /api/presentations/generate  → Generate New Presentation
@presentations_bp.route('/generate', methods=['POST'])
@token_required
def generate_presentation(user_id):
    """
    Generate a new presentation using AI with theme and text amount support.
    ✅ UPDATED: Now accepts theme and text_amount from frontend
    ✅ UPDATED: Auth integrated (user_id from token)
    """
    try:
        data = request.get_json() or {}

        # STEP 1: VALIDATE PROMPT
        prompt = (data.get('prompt') or '').strip()
        if not prompt:
            return jsonify({'success': False, 'error': 'Prompt is required'}), 400

        # STEP 2: DETERMINE SLIDE COUNT
        raw_count = data.get('slides_count') or data.get('slides') or 8
        try:
            slides_count = int(raw_count)
        except: 
            slides_count = 8

        slides_count = max(3, min(slides_count, 20))

        # Custom Outline Override
        custom_outline = data.get('custom_outline')
        if custom_outline: 
            lines = [line for line in custom_outline.split('\n') if line.strip()]
            if len(lines) > 0:
                slides_count = len(lines)
                print(f"🔹 Custom Outline Detected: Overriding slides count to {slides_count}")

        tokens_needed = slides_count * TOKENS_PER_SLIDE
        quota_reservation = reserve_daily_tokens(user_id, tokens_needed)
        if not quota_reservation['allowed']:
            return jsonify({
                'success': False,
                'error': quota_reservation['retry_message'],
                'quota': quota_reservation
            }), 429

        # STEP 3: CAPTURE ALL GENERATION OPTIONS
        ai_model = data.get('ai_model', 'gemini')
        image_source = data.get('image_source', 'real')
        style = data.get('style', 'professional')
        language = data.get('language', 'English')
        
        # ✅ THEME & TEXT AMOUNT
        theme = (data.get('theme') or 'dialogue').lower().strip()
        text_amount = (data.get('text_amount') or 'concise').lower().strip()
        
        image_style = data.get('image_style', 'photorealistic')
        use_search = bool(data.get('use_search', False))
        source_material = data.get('source_material')

        # user_id is now provided by @token_required decorator

        # STEP 4: LOG GENERATION REQUEST
        user_info = execute_query("SELECT name, email FROM users WHERE id = %s", (user_id,), fetch=True)
        user_display = f"User ID: {user_id}"
        if user_info:
            user_display = f"{user_info[0]['name']} ({user_info[0]['email']})"

        print(f"\n{'='*60}")
        print(f"🚀 [NEW GENERATION] Request from: {user_display}")
        print(f"   Prompt: '{prompt}'")
        print(f"   Slides: {slides_count} | Theme: {theme.upper()}")
        print(f"{'='*60}\n")
        
        # STEP 5: GENERATE SLIDES USING AI SERVICE
        if ai_service:
            slides = ai_service.generate_slides(
                prompt=prompt,
                slides_count=slides_count,
                custom_outline=custom_outline,
                style=style,
                language=language,
                theme=theme,              # ✅ PASS THEME
                image_source=image_source,
                text_amount=text_amount,  # ✅ PASS TEXT AMOUNT
                use_search=use_search,
                ai_model=ai_model,
                source_material=source_material
            )
        else:
            print("⚠️ AI service unavailable, using fallback content")
            outline_titles = []
            if custom_outline:
                outline_titles = [line.strip() for line in custom_outline.split('\n') if line.strip()]

            if outline_titles:
                slides = [
                    {
                        "title": outline_titles[i],
                        "content": "AI service is currently unavailable. Please check server logs.",
                        "image": None,
                        "layout": "hero"
                    }
                    for i in range(len(outline_titles))
                ]
            else:
                slides = [
                    {
                        "title": f"{prompt} Part {i + 1}",
                        "content": "AI service is currently unavailable. Please check server logs.",
                        "image": None,
                        "layout": "hero"
                    }
                    for i in range(slides_count)
                ]

        # STEP 6: SAVE TO DATABASE WITH THEME
        content_json = json.dumps(slides, ensure_ascii=False)

        pres_id = execute_query(
            """
            INSERT INTO presentations 
            (user_id, title, prompt, slides_count, content, theme, style, language, total_slides)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
            """,
            (user_id, prompt, prompt, slides_count, content_json, theme, style, language, slides_count)
        )

        # STEP 7: RETURN SUCCESS RESPONSE
        print(f"✅ Presentation {pres_id} created successfully")
        print(f"   Theme: {theme} | Text: {text_amount} | Slides: {len(slides)}")

        return jsonify({
            'success': True,
            'presentation_id': pres_id,
            'message': 'Presentation generated successfully',
            'presentation': {
                'id': pres_id,
                'title': prompt,
                'theme': theme,
                'text_amount': text_amount,
                'slides_count': len(slides)
            },
            'slides': slides,
            'quota': quota_reservation
        }), 201

    except Exception as e: 
        print(f"❌ Generate error: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@presentations_bp.route('/assistant-chat', methods=['POST'])
@token_required
def assistant_chat(user_id):
    """Chat assistant endpoint powered by Gemini/OpenRouter."""
    try:
        data = request.get_json() or {}
        question = (data.get('question') or '').strip()
        ai_model = (data.get('ai_model') or 'gemini').strip().lower()

        if not question:
            return jsonify({'success': False, 'error': 'Question is required'}), 400

        if not ai_service:
            return jsonify({'success': False, 'error': 'AI service unavailable'}), 503

        answer = ai_service.chat_assistant(question, ai_model=ai_model)
        if not answer:
            return jsonify({'success': False, 'error': 'Unable to generate assistant reply'}), 502

        return jsonify({'success': True, 'answer': answer, 'model': ai_model}), 200
    except Exception as e:
        print(f"❌ assistant_chat error: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

# DELETE /api/presentations/<id>  → Delete Single
@presentations_bp.route('/<int:pres_id>', methods=['DELETE'])
@token_required
def delete_presentation(user_id, pres_id):
    """Delete a single presentation"""
    try:
        execute_query("DELETE FROM presentations WHERE id = %s AND user_id = %s", (pres_id, user_id))
        print(f"🗑️ Presentation {pres_id} deleted by User {user_id}")
        return jsonify({'success': True, 'message': 'Presentation deleted'}), 200
    except Exception as e:
        print(f"❌ delete_presentation error: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

# DELETE /api/presentations/all  → Delete ALL
@presentations_bp.route('/all', methods=['DELETE'])
@token_required
def delete_all_presentations(user_id):
    """Delete all presentations for current user"""
    try:
        execute_query("DELETE FROM presentations WHERE user_id = %s", (user_id,))
        print(f"🗑️ All presentations deleted for User {user_id}")
        return jsonify({'success': True, 'message': 'All presentations deleted'}), 200
    except Exception as e: 
        print(f"❌ delete_all_presentations error: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

# POST /api/presentations/bulk-delete  → Bulk Delete Multiple
@presentations_bp.route('/bulk-delete', methods=['POST'])
@token_required
def bulk_delete_presentations(user_id):
    """
    Bulk delete presentations
    ✅ UPDATED: Verify ownership
    """
    try:
        data = request.get_json()
        
        presentation_ids = data.get('presentation_ids', [])
        
        if not presentation_ids or not isinstance(presentation_ids, list):
            return jsonify({
                "success": False,
                "error": "Invalid presentation_ids format"
            }), 400
        
        if len(presentation_ids) == 0:
            return jsonify({
                "success": False,
                "error": "No presentations selected"
            }), 400
        
        # Safety limit
        if len(presentation_ids) > 100:
            return jsonify({
                "success": False,
                "error": "Maximum 100 presentations can be deleted at once"
            }), 400
        
        print(f"🗑️ Bulk delete requested by User {user_id}: {len(presentation_ids)} items")
        
        deleted_count = 0
        failed_count = 0
        errors = []
        
        for presentation_id in presentation_ids:
            try:
                # Check if exists and belongs to user
                rows = execute_query(
                    "SELECT id FROM presentations WHERE id = %s AND user_id = %s",
                    (presentation_id, user_id),
                    fetch=True
                )
                
                if not rows:
                    print(f"   ⚠️ Presentation {presentation_id} not found or Access Denied")
                    failed_count += 1
                    errors.append(f"Presentation {presentation_id}: Not found or access denied")
                    continue
                
                # Delete
                execute_query(
                    "DELETE FROM presentations WHERE id = %s AND user_id = %s",
                    (presentation_id, user_id)
                )
                
                deleted_count += 1
                print(f"   ✅ Deleted: {presentation_id}")
                
            except Exception as e:
                print(f"   ❌ Failed to delete {presentation_id}: {e}")
                failed_count += 1
                errors.append(f"Presentation {presentation_id}: {str(e)}")
        
        print(f"✅ Bulk delete completed: {deleted_count} deleted, {failed_count} failed")
        
        response = {
            "success": True,
            "deleted_count": deleted_count,
            "failed_count": failed_count,
            "total_requested": len(presentation_ids)
        }
        
        if errors:
            response["errors"] = errors
        
        return jsonify(response), 200
        
    except Exception as e:
        print(f"❌ Bulk delete error: {e}")
        traceback.print_exc()
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500

# EXPORT HANDLER (PPTX/PDF) - ✅ THEME-AWARE VERSION
def handle_export(pres_id, format_type, user_id):
    """
    Export presentation to PPTX/PDF/DOCX/DOC with correct theme
    ✅ FIXED: Now properly passes theme from database
    """
    try:
        # Normalize format input to avoid case sensitivity issues
        format_type = str(format_type).lower().strip()

        print(f"📤 [EXPORT] Starting export for presentation {pres_id} as {format_type.upper()} for user {user_id}")
        
        # ⭐ FETCH PRESENTATION WITH THEME AND USER CHECK
        rows = execute_query(
            "SELECT id, title, prompt, content, theme FROM presentations WHERE id = %s AND user_id = %s",
            (pres_id, user_id),
            fetch=True
        )
        
        if not rows:
            print(f"❌ [EXPORT] Presentation {pres_id} not found")
            return jsonify({'success': False, 'error': 'Not found'}), 404

        pres = rows[0]
        
        # ⭐ GET THEME FROM DATABASE
        theme = pres.get('theme', 'dialogue')
        if not theme or theme.strip() == '':
            theme = 'dialogue'
        
        theme = theme.lower().strip()
        
        print(f"🎨 [EXPORT] Database theme: '{theme}'")

        # Parse slides from JSON
        try:
            raw = pres.get('content')
            slides_data = json.loads(raw) if isinstance(raw, str) and raw.strip() else (raw or [])
        except Exception as e:
            print(f"⚠️ [EXPORT] JSON parse error: {e}")
            slides_data = []

        print(f"📊 [EXPORT] Loaded {len(slides_data)} slides")

        # ✅ ADAPTER OBJECT WITH THEME
        class PresentationData:
            def __init__(self, id, title, prompt, slides, theme):
                self.id = id
                self.title = title
                self.prompt = prompt
                self.content = {'slides': slides}
                self.theme = theme  # ⭐ CRITICAL: Pass theme here
        
        pres_data_obj = PresentationData(
            pres['id'],
            pres['title'],
            pres['prompt'],
            slides_data,
            theme  # ⭐ PASS THEME FROM DATABASE
        )
        
        print(f"✅ [EXPORT] PresentationData created with theme: '{pres_data_obj.theme}'")

        # Select service and set file properties
        if format_type == 'pptx':
            service = PPTXService()
            mimetype = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
            ext = 'pptx'
        elif format_type == 'pdf':
            if PDFService is None:
                return jsonify({'success': False, 'error': 'PDF service not available'}), 500
            service = PDFService()
            mimetype = 'application/pdf'
            ext = 'pdf'
        elif format_type == 'docx':
            if DOCXService is None:
                return jsonify({'success': False, 'error': 'DOCX service not available'}), 500
            service = DOCXService()
            mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            ext = 'docx'
        elif format_type == 'doc':
            if DOCService is None:
                return jsonify({'success': False, 'error': 'DOC service not available'}), 500
            service = DOCService()
            mimetype = 'application/msword'
            ext = 'doc'
        else:
            return jsonify({'success': False, 'error': "Invalid format. Use 'pptx', 'pdf', 'docx', or 'doc'"}), 400

        # ✅ Generate file WITH THEME
        print(f"🔧 [EXPORT] Calling {format_type.upper()} service with theme: {theme}")
        file_content = service.generate(pres_data_obj)

        print(f"✅ [EXPORT] {format_type.upper()} generated successfully with theme: {theme}")

        return send_file(
            io.BytesIO(file_content),
            mimetype=mimetype,
            as_attachment=True,
            download_name=f"{pres['title']}.{ext}"
        )

    except Exception as e:
        print(f"❌ [EXPORT] Error: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

# EXPORT ROUTES (Multiple URL patterns for compatibility)

@presentations_bp.route('/<int:pres_id>/download/<format>', methods=['GET'])
@token_required
def download_presentation(user_id, pres_id, format):
    """
    Modern route: /api/presentations/123/download/pptx
    """
    return handle_export(pres_id, format, user_id)

@presentations_bp.route('/<int:pres_id>/export/pptx', methods=['GET'])
@token_required
def export_pptx_legacy(user_id, pres_id):
    """
    Legacy route: /api/presentations/123/export/pptx
    """
    return handle_export(pres_id, 'pptx', user_id)

@presentations_bp.route('/<int:pres_id>/export/pdf', methods=['GET'])
@token_required
def export_pdf_legacy(user_id, pres_id):
    """
    Legacy route: /api/presentations/123/export/pdf
    """
    return handle_export(pres_id, 'pdf', user_id)

@presentations_bp.route('/<int:pres_id>/export/docx', methods=['GET'])
@token_required
def export_docx_legacy(user_id, pres_id):
    """
    Legacy route: /api/presentations/123/export/docx
    """
    return handle_export(pres_id, 'docx', user_id)

@presentations_bp.route('/<int:pres_id>/export/doc', methods=['GET'])
@token_required
def export_doc_legacy(user_id, pres_id):
    """
    Legacy route: /api/presentations/123/export/doc
    """
    return handle_export(pres_id, 'doc', user_id)

# EXPORT WITH QUERY PARAM (Optional alternative)
@presentations_bp.route('/<int:pres_id>/export', methods=['GET'])
@token_required
def export_with_query(user_id, pres_id):
    """
    Query param route: /api/presentations/123/export?format=pptx
    """
    format_type = request.args.get('format', 'pptx').lower()
    return handle_export(pres_id, format_type, user_id)

# DEBUG ROUTE (Optional - for testing)
@presentations_bp.route('/<int:pres_id>/debug', methods=['GET'])
@token_required
def debug_presentation(user_id, pres_id):
    """
    Debug endpoint to see raw database content including theme
    """
    try:
        rows = execute_query(
            "SELECT * FROM presentations WHERE id = %s AND user_id = %s",
            (pres_id, user_id),
            fetch=True
        )

        if not rows:
            return jsonify({'error': 'Not found'}), 404

        pres = rows[0]

        # Parse content
        try:
            raw = pres.get('content')
            if isinstance(raw, str):
                pres['content_parsed'] = json.loads(raw)
            else:
                pres['content_parsed'] = raw
        except: 
            pres['content_parsed'] = None

        # Format dates
        for key in ['created_at', 'updated_at']:
            if isinstance(pres.get(key), (datetime.datetime, datetime.date)):
                pres[key] = pres[key].isoformat()

        # ⭐ HIGHLIGHT THEME
        print(f"🔍 [DEBUG] Presentation {pres_id} theme: '{pres.get('theme', 'NOT SET')}'")

        return jsonify({'success': True, 'data': pres}), 200

    except Exception as e:
        return jsonify({'error': str(e)}), 500

# STATISTICS ROUTE (Optional - for analytics)
@presentations_bp.route('/stats', methods=['GET'])
@token_required
def get_statistics(user_id):
    """
    Get presentation statistics
    """
    try:
        # Total presentations
        total_result = execute_query(
            "SELECT COUNT(*) as total FROM presentations WHERE user_id = %s",
            (user_id,),
            fetch=True
        )
        total = total_result[0]['total'] if total_result else 0

        # By theme
        theme_stats = execute_query(
            """
            SELECT theme, COUNT(*) as count 
            FROM presentations 
            WHERE user_id = %s
            GROUP BY theme 
            ORDER BY count DESC
            """,
            (user_id,),
            fetch=True
        ) or []

        # By style
        style_stats = execute_query(
            """
            SELECT style, COUNT(*) as count 
            FROM presentations 
            WHERE user_id = %s
            GROUP BY style 
            ORDER BY count DESC
            """,
            (user_id,),
            fetch=True
        ) or []

        # Recent activity
        recent = execute_query(
            """
            SELECT DATE(created_at) as date, COUNT(*) as count 
            FROM presentations 
            WHERE user_id = %s AND created_at >= (NOW() - INTERVAL '30 days')
            GROUP BY DATE(created_at)
            ORDER BY date DESC
            """,
            (user_id,),
            fetch=True
        ) or []

        return jsonify({
            'success': True,
            'stats': {
                'total': total,
                'by_theme': theme_stats,
                'by_style': style_stats,
                'recent_activity': recent
            }
        }), 200

    except Exception as e: 
        print(f"❌ Stats error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

# STARTUP LOG
print("✅ All presentation routes registered:")
print("   GET    /api/presentations/")
print("   GET    /api/presentations/<id>")
print("   PUT    /api/presentations/<id>")
print("   PUT    /api/presentations/<id>/theme  ⭐ NEW")
print("   POST   /api/presentations/generate")
print("   DELETE /api/presentations/<id>")
print("   DELETE /api/presentations/all")
print("   POST   /api/presentations/bulk-delete")
print("   GET    /api/presentations/<id>/download/<format>")
print("   GET    /api/presentations/<id>/export/pptx")
print("   GET    /api/presentations/<id>/export/pdf")
print("   GET    /api/presentations/<id>/export?format=pptx")
print("   GET    /api/presentations/<id>/debug")
print("   GET    /api/presentations/stats")