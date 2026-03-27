import os

os.environ['GRPC_DNS_RESOLVER'] = 'native'

import sys
from dotenv import load_dotenv
 

# LOAD ENVIRONMENT VARIABLES
print("\n" + "="*80)
print("🔧 GAMMA AI - INITIALIZATION")
print("="*80)

# Get project root directory
project_root = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(project_root, '.env')

print(f"📂 Project Root: {project_root}")
print(f"🔍 Loading .env from: {env_path}")

if os.path.exists(env_path):
    load_dotenv(env_path)
    print("✅ .env file loaded successfully")
    
    # Verify critical environment variables
    api_key = os.getenv('GOOGLE_GEMINI_API_KEY')
    if api_key:
        print(f"✅ GOOGLE_GEMINI_API_KEY: {api_key[:10]}... (Hidden)")
    else:
        print("⚠️  GOOGLE_GEMINI_API_KEY not found in .env (will check fallback)")
    
    database_url = os.getenv('DATABASE_URL', '').strip()
    if database_url:
        print("✅ Database: PostgreSQL via DATABASE_URL")
    else:
        db_host = os.getenv('DB_HOST', 'localhost')
        db_name = os.getenv('DB_NAME', 'gamma_ai')
        print(f"✅ Database fallback: {db_host}/{db_name}")
else:
    print(f"⚠️  .env file not found at: {env_path}")
    print("   Using default configuration")

print("="*80 + "\n")

# IMPORT AND CREATE FLASK APP
try:
    from app import create_app

    print("🚀 Creating Flask application...")
    app = create_app()
    print("✅ Flask app created successfully\n")

except ImportError as e:
    print(f"❌ Import error: {e}")
    print("\n💡 Make sure you have these folders:")
    print("   - app/")
    print("   - app/__init__.py")
    print("   - app/routes/")
    print("   - app/models/")
    print("   - app/services/")
    sys.exit(1)

except Exception as e:
    print(f"❌ Failed to create Flask app: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# MAIN ENTRY POINT
if __name__ == '__main__':
    print("\n" + "="*80)
    print("🎨 GAMMA AI - Python Flask Edition")
    print("="*80)
    print("🚀 Server starting on http://localhost:5000")
    print("📊 Database: PostgreSQL (Neon compatible)")
    print("🤖 AI: Google Gemini 2.0 Flash")
    print("📄 Export: PDF/DOCX/PPTX")
    print("👤 Author: GuptaSigma")
    print("📅 Date: 2025-11-23")
    print("="*80)
    print("\n💡 Access the app:")
    print("   🌐 Local:   http://localhost:5000")
    print("\n💡 Press CTRL+C to stop the server\n")
    
    try:
        port = int(os.getenv('PORT', '5000'))
        debug_enabled = os.getenv('FLASK_DEBUG', '0').lower() in ('1', 'true', 'yes')

        app.run(
            host='0.0.0.0',
            port=port,
            debug=debug_enabled,
            use_reloader=debug_enabled
        )
    except KeyboardInterrupt:
        print("\n\n" + "="*80)
        print("👋 Server stopped by user")
        print("="*80 + "\n")
    except Exception as e:
        print(f"\n❌ Server error: {e}\n")
        import traceback
        traceback.print_exc()
        sys.exit(1)