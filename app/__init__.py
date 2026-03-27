import os
from pathlib import Path

try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None
from flask import Flask
from flask_cors import CORS

def create_app():
    """Create and configure Flask application"""

    if load_dotenv:
        root_dir = Path(__file__).resolve().parents[1]
        env_path = root_dir / '.env'
        if env_path.exists():
            load_dotenv(env_path)
    
    app = Flask(__name__, 
                template_folder='../templates',
                static_folder='../static')
    
    # Configuration — values MUST come from .env (no hardcoded fallbacks)
    secret_key = os.getenv('SECRET_KEY')
    jwt_secret_key = os.getenv('JWT_SECRET_KEY')

    if not secret_key:
        raise RuntimeError("❌ SECRET_KEY missing from .env file!")
    if not jwt_secret_key:
        raise RuntimeError("❌ JWT_SECRET_KEY missing from .env file!")

    app.config['SECRET_KEY'] = secret_key
    app.config['JWT_SECRET_KEY'] = jwt_secret_key
    
    # Enable CORS
    CORS(app, resources={r"/api/*": {"origins": "*"}})
    
    # Register blueprints
    with app.app_context():
        try:
            from app.routes.main import main_bp
            from app.routes.auth import auth_bp
            from app.routes.presentations import presentations_bp
            
            app.register_blueprint(main_bp)
            app.register_blueprint(auth_bp, url_prefix='/api/auth')
            app.register_blueprint(presentations_bp, url_prefix='/api/presentations')
            
            print("✅ All blueprints registered successfully")
            
            # Initialize database
            from app.models.database import init_db
            init_db()
            
        except Exception as e:
            print(f"❌ Blueprint registration error: {e}")
            import traceback
            traceback.print_exc()
            raise
    
    return app