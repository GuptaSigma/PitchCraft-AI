# 🎨 PitchCraft AI - AI-Powered Presentation Generator

A powerful Flask-based web application that generates professional presentations using AI. Create stunning slideshows with real images from Wikipedia, multiple AI models (Gemini, ChatGPT, DeepSeek), and 8 beautiful themes.

---

## ✨ Features

- **🤖 Multiple AI Models**
  - Google Gemini 3 Flash (Primary)
  - ChatGPT (via OpenRouter)
  - DeepSeek R1 (via OpenRouter)
  
- **🖼️ Real Image Integration**
  - Wikipedia API (Primary - High-quality real images)
  - Google Custom Search (Optional - Requires permissions)
  - ClipDrop AI (Fallback)
  
- **🎨 8 Professional Themes**
  - Dialogue (Blue gradient)
  - Alien (Green tech)
  - Wine (Purple elegant)
  - Snowball (Clean white)
  - Petrol (Teal modern)
  - Piano (Classic black/white)
  - Sunset (Orange warm)
  - Midnight (Dark blue)
  
- **📐 8 Flexible Layouts**
  - Centered, Fixed Information, Three Column
  - Split Sidebar, Image Only, Big Header
  - Content Focus, Image Left

- **📤 Export Formats**
  - PowerPoint (.pptx)
  - PDF (.pdf)
  - Word Document (.docx)

- **🔐 User Authentication**
  - JWT-based secure login
  - bcrypt password hashing
  - Session management

---

## 🛠️ Tech Stack

### Backend
- **Python 3.x** - Core language
- **Flask** - Web framework
- **MySQL 8.x** - Database with connection pooling
- **JWT** - Authentication tokens
- **bcrypt** - Password security

### Frontend
- **HTML5/CSS3/JavaScript**
- **Tailwind CSS** - Utility-first styling
- **Fetch API** - Async requests

### AI & APIs
- **Google Gemini API** - AI content generation
- **OpenRouter API** - ChatGPT & DeepSeek access
- **Wikipedia API** - Real image fetching
- **Google Custom Search** - Image search (optional)
- **ClipDrop API** - AI image generation (fallback)

### Python Libraries
```
flask
mysql-connector-python
PyJWT
bcrypt
requests
python-pptx
reportlab
python-docx
Pillow
python-dotenv
```

---

## 📥 Installation

### 1. Clone Repository
```bash
cd "C:\Users\POOJA GUPTA\Pictures\doc"
```

### 2. Create Virtual Environment
```powershell
python -m venv .venv
.venv\Scripts\activate
```

### 3. Install Dependencies
```powershell
pip install -r requirements.txt
```

### 4. Setup MySQL Database
```sql
CREATE DATABASE gamma_ai;
USE gamma_ai;

CREATE TABLE users (
    id INT AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(255) UNIQUE NOT NULL,
    email VARCHAR(255) UNIQUE NOT NULL,
    password VARCHAR(255) NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE presentations (
    id INT AUTO_INCREMENT PRIMARY KEY,
    user_id INT NOT NULL,
    title VARCHAR(255) NOT NULL,
    slides JSON NOT NULL,
    theme VARCHAR(50) DEFAULT 'dialogue',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (user_id) REFERENCES users(id)
);
```

### 5. Configure Environment Variables

Create `.env` file:
```env
# Database Configuration
DB_HOST=localhost
DB_USER=root
DB_PASSWORD=your_mysql_password
DB_NAME=gamma_ai
DB_PORT=3306

# Flask Secrets
SECRET_KEY=your-secret-key-here
JWT_SECRET_KEY=your-jwt-secret-key-here
FLASK_ENV=development
FLASK_RUN_PORT=5000

# AI API Keys
GOOGLE_GEMINI_API_KEY=your_gemini_api_key_here
OPENROUTER_API_KEY=your_openrouter_api_key_here
CLIPDROP_API_KEY=your_clipdrop_api_key_here

# Google Custom Search (Optional - requires Google Cloud access)
GOOGLE_CSE_API_KEY=your_google_cse_key_here
GOOGLE_CSE_CX=your_custom_search_engine_id

# File Paths
UPLOAD_FOLDER=./uploads
STATIC_GENERATED_DIR=./app/static/generated
IMAGE_DOWNLOAD_TIMEOUT=12
```

---

## 🔑 API Setup Guide

### 1. Google Gemini API
1. Go to [Google AI Studio](https://makersuite.google.com/app/apikey)
2. Create API key
3. Add to `.env`: `GOOGLE_GEMINI_API_KEY=AIzaSy...`

### 2. OpenRouter (ChatGPT/DeepSeek)
1. Go to [OpenRouter](https://openrouter.ai/)
2. Sign up and create API key
3. Add credits to account
4. Add to `.env`: `OPENROUTER_API_KEY=sk-or-v1-...`

### 3. ClipDrop (Optional - Image Fallback)
1. Go to [ClipDrop](https://clipdrop.co/apis)
2. Create API key
3. Add to `.env`: `CLIPDROP_API_KEY=...`

### 4. Wikipedia API
**No setup required!** Wikipedia API is free and works out of the box.

### 5. Google Custom Search (Optional)
⚠️ **Currently disabled** due to permissions. If you have Google Cloud access:
1. Create new Google Cloud project
2. Enable Custom Search API
3. Create Search Engine at [Programmable Search](https://programmablesearchengine.google.com/)
4. Add credentials to `.env`

---

## 🚀 Running the Application

### Start Server
```powershell
python run.py
```

### Access Application
Open browser: **http://localhost:5000**

### Create Account
1. Click **Sign Up**
2. Enter username, email, password
3. Login with credentials

### Generate Presentation
1. Enter topic (e.g., "Climate Change")
2. Select AI model (Gemini/ChatGPT/DeepSeek)
3. Choose theme (8 options)
4. Set number of slides (1-15)
5. Click **Generate**

### Export
- Click Export button
- Choose format: PPTX, PDF, or DOCX
- Download automatically starts

---

## 📁 Project Structure

```
doc/
├── app/
│   ├── __init__.py              # Flask app factory
│   ├── models/
│   │   ├── database.PY          # MySQL connection pool
│   │   ├── user.py              # User model
│   │   └── presentation.py      # Presentation model
│   ├── routes/
│   │   ├── auth.py              # Login/signup/JWT
│   │   ├── main.py              # Homepage routes
│   │   └── presentations.py     # CRUD operations
│   └── services/
│       ├── ai_service.py        # AI integration (Gemini/OpenRouter/Wikipedia)
│       ├── pptx_service.py      # PowerPoint generation
│       ├── pdf_services.py      # PDF export
│       └── docx_services.py     # Word export
├── static/
│   ├── css/style.css            # Global styles
│   ├── js/
│   │   ├── auth.js              # Authentication logic
│   │   ├── api.js               # API calls
│   │   └── editor.js            # Slide editor
│   └── generated/               # Exported files
├── templates/
│   ├── index.html               # Landing page
│   ├── login.html               # Login form
│   ├── signup.html              # Registration
│   ├── dashboard.html           # User dashboard
│   ├── presentation.html        # Slide editor
│   └── view_presentation.html   # Read-only view
├── uploads/                     # User uploads
├── .env                         # Environment variables (gitignored)
├── .gitignore                   # Git exclusions
├── requirements.txt             # Python dependencies
├── run.py                       # App entry point
└── README.md                    # This file
```

---

## 🎨 Theme System

All themes support **smart text colors** based on background brightness:

| Theme | Colors | Best For |
|-------|--------|----------|
| **Dialogue** | Blue gradient | Corporate, Tech |
| **Alien** | Green tech | Science, Innovation |
| **Wine** | Purple elegant | Luxury, Creative |
| **Snowball** | Clean white | Minimal, Professional |
| **Petrol** | Teal modern | Modern, Fresh |
| **Piano** | Black/white | Classic, Bold |
| **Sunset** | Orange warm | Energy, Passion |
| **Midnight** | Dark blue | Elegant, Serious |

---

## 🖼️ Image Fetching Strategy

### Current Setup (Wikipedia Primary)
```
1. Wikipedia API (PRIMARY)
   ✅ Free, no authentication
   ✅ High-quality real images
   ✅ Reliable

2. ClipDrop AI (FALLBACK)
   🎨 AI-generated images
   💰 Requires API key
```

### Optional (Google Custom Search)
```
Disabled due to project permissions
Requires Google Cloud access
Can be re-enabled in ai_service.py
```

---

## 🔧 Configuration

### Update AI Model
Edit [app/services/ai_service.py](app/services/ai_service.py):
```python
def _get_ai_text(self, topic, slides_count, model):
    if model in ["gemini", "gemini-flash"]:
        return self._call_gemini(prompt)
    elif model in ["chatgpt", "gpt", "openai"]:
        return self._call_openrouter_chatgpt(prompt)
    elif model in ["deepseek"]:
        return self._call_deepseek(prompt)
```

### Add New Theme
Edit [static/css/style.css](static/css/style.css):
```css
[data-theme="new-theme"] {
    --primary-color: #your-color;
    --bg-color: #your-bg;
    --text-color: #your-text;
}
```

---

## 🐛 Troubleshooting

### Problem: Server won't start
**Solution:** Check MySQL is running and credentials in `.env` are correct
```powershell
mysql -u root -p
SHOW DATABASES;
```

### Problem: AI not generating content
**Solution:** Verify API keys are valid and have credit
```powershell
# Check .env file
cat .env | Select-String "API_KEY"
```

### Problem: Images not loading
**Solution:** Wikipedia is primary source, no config needed. If you see errors:
- Check internet connection
- Verify topic has Wikipedia page
- Check ClipDrop API key if Wikipedia fails

### Problem: 401/403 API Errors
**Solution:** 
- **OpenRouter:** Check billing/credits at https://openrouter.ai/
- **Gemini:** Verify key at https://makersuite.google.com/app/apikey
- **Google CSE:** Currently disabled (requires Cloud project access)

### Problem: Export fails
**Solution:** Check `static/generated/` directory exists and is writable
```powershell
mkdir -Force static/generated
```

---

## 🔒 Security Notes

⚠️ **IMPORTANT:**
1. **Never commit `.env` file** (already in `.gitignore`)
2. **Rotate API keys regularly**
3. **Use strong JWT secrets** in production
4. **Enable HTTPS** for production deployment
5. **Sanitize user inputs** (already implemented)

---

## 📝 API Endpoints

### Authentication
- `POST /signup` - Create account
- `POST /login` - Get JWT token
- `GET /verify` - Verify token
- `POST /logout` - Invalidate token

### Presentations
- `GET /dashboard` - List user presentations
- `POST /generate` - Create new presentation
- `GET /presentations/<id>` - View/edit presentation
- `PUT /presentations/<id>` - Update slides
- `DELETE /presentations/<id>` - Delete presentation
- `POST /export/<format>` - Export (pptx/pdf/docx)

---

## 🎯 Roadmap

- [ ] Add Google Slides export
- [ ] Real-time collaboration
- [ ] Template marketplace
- [ ] Voice narration
- [ ] Video export
- [ ] Mobile app
- [ ] Custom fonts
- [ ] Animation presets

---

## 👨‍💻 Developer

**Pooja Gupta**  
PitchCraft AI - Gamma AI Project

---

## 📄 License

This project is for educational/portfolio purposes.

---

## 🙏 Acknowledgments

- Google Gemini API
- OpenRouter (ChatGPT/DeepSeek access)
- Wikipedia Foundation
- Flask Framework
- Tailwind CSS

---

**Made with ❤️ using Python & AI**
