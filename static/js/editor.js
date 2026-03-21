/**
 * Editor Logic - COMPLETE FIXED VERSION
 * Gamma AI - Python Flask Edition
 */

// ==========================================
// GLOBAL STATE
// ==========================================
let currentPresentation = null;
let selectedTheme = 'alien';
let selectedImageStyle = 'illustration';

console.log('🎨 Editor initialized with defaults:', {
    theme: selectedTheme,
    imageStyle: selectedImageStyle
});

// ==========================================
// THEME SELECTION
// ==========================================

/**
 * Select theme
 */
function selectTheme(element, theme) {
    console.log('🎨 Theme selection clicked:', theme);
    
    // Remove selected class from all theme options
    document.querySelectorAll('.theme-option').forEach(el => {
        el.classList.remove('selected');
    });
    
    // Add selected class to clicked element
    element.classList.add('selected');
    
    // Update global state
    selectedTheme = theme;
    
    console.log('✅ Theme selected:', selectedTheme);
}

/**
 * Select image style
 */
function selectImageStyle(element, style) {
    console.log('🖼️ Image style selection clicked:', style);
    
    // Remove selected class from all image style options
    document.querySelectorAll('.image-style-option').forEach(el => {
        el.classList.remove('selected');
    });
    
    // Add selected class to clicked element
    element.classList.add('selected');
    
    // Update global state
    selectedImageStyle = style;
    
    console.log('✅ Image style selected:', selectedImageStyle);
}

// ==========================================
// PRESENTATION GENERATION
// ==========================================

/**
 * Generate presentation - COMPLETE FIXED VERSION
 */
async function generatePresentation() {
    console.log('🚀 Generate presentation called');
    console.log('📊 Current state:', {
        theme: selectedTheme,
        imageStyle: selectedImageStyle
    });
    
    const promptInput = document.getElementById('promptInput');
    const slidesCount = document.getElementById('slidesCount').value;
    const style = document.getElementById('styleSelect').value;
    const language = document.getElementById('languageSelect').value;
    const generateBtn = document.getElementById('generateBtn');
    
    // Validate prompt
    if (!promptInput || !promptInput.value.trim()) {
        showToast('⚠️ Please enter a topic!', 'error');
        if (promptInput) promptInput.focus();
        return;
    }
    
    // Check authentication
    const token = localStorage.getItem('token');
    if (!token) {
        showToast('❌ Please login first!', 'error');
        setTimeout(() => window.location.href = '/', 1500);
        return;
    }
    
    console.log('✅ Token found');
    
    // Disable button
    generateBtn.disabled = true;
    generateBtn.textContent = '⏳ Generating...';
    
    // Show loading
    const slidesPreview = document.getElementById('slidesPreview');
    if (slidesPreview) {
        slidesPreview.innerHTML = `
            <div class="loading-state">
                <div class="loading-spinner"></div>
                <h3>🤖 AI is creating your presentation...</h3>
                <p>This may take a few moments</p>
            </div>
        `;
    }
    
    try {
        // Prepare request data
        const requestData = {
            prompt: promptInput.value.trim(),
            slides_count: parseInt(slidesCount) || 10,
            style: style || 'business',
            language: language || 'en-uk',
            theme: selectedTheme || 'alien',
            image_style: selectedImageStyle || 'illustration'
        };
        
        console.log('📤 Sending request:', requestData);
        
        // Make request
        const API_BASE = window.location.origin + '/api';
        const response = await fetch(`${API_BASE}/presentations/generate`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}`
            },
            body: JSON.stringify(requestData)
        });
        
        console.log('📥 Response status:', response.status);
        
        // Parse response
        const data = await response.json();
        console.log('📦 Response data:', data);
        
        if (!response.ok) {
            throw new Error(data.error || `Request failed with status ${response.status}`);
        }
        
        if (data.success) {
            currentPresentation = data.content;
            renderSlides(data.content);
            showToast('✅ Presentation generated successfully!', 'success');
            
            // Show export buttons
            const exportButtons = document.getElementById('exportButtons');
            if (exportButtons) {
                exportButtons.style.display = 'flex';
            }
        } else {
            throw new Error(data.message || 'Generation failed');
        }
        
    } catch (error) {
        console.error('❌ Generation error:', error);
        showToast('❌ ' + error.message, 'error');
        
        if (slidesPreview) {
            slidesPreview.innerHTML = `
                <div class="empty-state">
                    <div class="empty-icon">❌</div>
                    <h3>Generation Failed</h3>
                    <p>${error.message}</p>
                    <button class="btn btn-primary" onclick="generatePresentation()">Try Again</button>
                </div>
            `;
        }
    } finally {
        generateBtn.disabled = false;
        generateBtn.textContent = '🚀 Generate Presentation';
    }
}

// ==========================================
// RENDER SLIDES
// ==========================================

/**
 * Render slides in preview
 */
function renderSlides(content) {
    const slidesPreview = document.getElementById('slidesPreview');
    const contentTitle = document.getElementById('contentTitle');
    
    if (!slidesPreview || !content) return;
    
    if (contentTitle) {
        contentTitle.textContent = content.title || 'Presentation';
    }
    
    let html = `
        <div class="slides-header">
            <div class="slides-info">
                <h3>${content.title || 'Untitled Presentation'}</h3>
                <p>${content.total_slides || 0} slides • ${content.theme_name || 'Default'} theme • ${content.style_name || 'Business'} style</p>
            </div>
        </div>
        <div class="slides-list">
    `;
    
    const slides = content.slides || [];
    
    slides.forEach((slide, index) => {
        html += `
            <div class="slide-card" data-slide-id="${index}">
                <div class="slide-header">
                    <div class="slide-number">Slide ${index + 1}</div>
                </div>
                <h3 class="slide-title">${slide.title || `Slide ${index + 1}`}</h3>
                <div class="slide-content">${formatSlideContent(slide.content || '')}</div>
                ${slide.image ? `<div class="slide-image-preview" style="background-image: url('${slide.image}');"></div>` : ''}
            </div>
        `;
    });
    
    html += '</div>';
    
    slidesPreview.innerHTML = html;
    
    console.log('✅ Rendered', slides.length, 'slides');
}

/**
 * Format slide content for display
 */
function formatSlideContent(content) {
    if (!content) return '';
    
    // Convert markdown-style to HTML
    content = content.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    content = content.replace(/\n/g, '<br>');
    
    // Limit preview length
    const maxLength = 500;
    if (content.length > maxLength) {
        content = content.substring(0, maxLength) + '...';
    }
    
    return content;
}

// ==========================================
// EXPORT FUNCTIONS
// ==========================================

/**
 * Export to PDF
 */
async function exportPDF() {
    if (!currentPresentation || !currentPresentation.id) {
        showToast('⚠️ No presentation to export', 'error');
        return;
    }
    
    showToast('📄 Generating PDF...', 'info');
    
    try {
        const token = localStorage.getItem('token');
        const API_BASE = window.location.origin + '/api';
        
        const response = await fetch(`${API_BASE}/presentations/${currentPresentation.id}/pdf`, {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });
        
        if (!response.ok) throw new Error('PDF export failed');
        
        const blob = await response.blob();
        downloadBlob(blob, `${currentPresentation.title}.pdf`);
        showToast('✅ PDF downloaded successfully!', 'success');
    } catch (error) {
        showToast('❌ PDF export failed: ' + error.message, 'error');
    }
}

/**
 * Export to DOCX
 */
async function exportDOCX() {
    if (!currentPresentation || !currentPresentation.id) {
        showToast('⚠️ No presentation to export', 'error');
        return;
    }
    
    showToast('📝 Generating DOCX...', 'info');
    
    try {
        const token = localStorage.getItem('token');
        const API_BASE = window.location.origin + '/api';
        
        const response = await fetch(`${API_BASE}/presentations/${currentPresentation.id}/docx`, {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });
        
        if (!response.ok) throw new Error('DOCX export failed');
        
        const blob = await response.blob();
        downloadBlob(blob, `${currentPresentation.title}.docx`);
        showToast('✅ DOCX downloaded successfully!', 'success');
    } catch (error) {
        showToast('❌ DOCX export failed: ' + error.message, 'error');
    }
}

/**
 * Export to PPTX
 */
async function exportPPTX() {
    if (!currentPresentation || !currentPresentation.id) {
        showToast('⚠️ No presentation to export', 'error');
        return;
    }
    
    showToast('📊 Generating PPTX...', 'info');
    
    try {
        const token = localStorage.getItem('token');
        const API_BASE = window.location.origin + '/api';
        
        const response = await fetch(`${API_BASE}/presentations/${currentPresentation.id}/pptx`, {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });
        
        if (!response.ok) throw new Error('PPTX export failed');
        
        const blob = await response.blob();
        downloadBlob(blob, `${currentPresentation.title}.pptx`);
        showToast('✅ PPTX downloaded successfully!', 'success');
    } catch (error) {
        showToast('❌ PPTX export failed: ' + error.message, 'error');
    }
}

/**
 * Download blob helper
 */
function downloadBlob(blob, filename) {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.style.display = 'none';
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);
}

// ==========================================
// STYLES
// ==========================================

// Add loading spinner styles
if (!document.getElementById('loading-styles')) {
    const style = document.createElement('style');
    style.id = 'loading-styles';
    style.textContent = `
        .loading-state {
            text-align: center;
            padding: 80px 20px;
        }
        
        .loading-spinner {
            width: 60px;
            height: 60px;
            border: 5px solid #e2e8f0;
            border-top-color: #667eea;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin: 0 auto 30px;
        }
        
        @keyframes spin {
            to { transform: rotate(360deg); }
        }
        
        .loading-state h3 {
            font-size: 24px;
            color: #1e293b;
            margin-bottom: 10px;
        }
        
        .loading-state p {
            color: #64748b;
            font-size: 16px;
        }
        
        .slides-header {
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 2px solid #e2e8f0;
        }
        
        .slides-info h3 {
            font-size: 28px;
            color: #1e293b;
            margin-bottom: 8px;
        }
        
        .slides-info p {
            color: #64748b;
            font-size: 15px;
        }
        
        .slide-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
        }
        
        .slide-number {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 13px;
            font-weight: 700;
        }
        
        .slide-image-preview {
            width: 100%;
            height: 200px;
            background-size: cover;
            background-position: center;
            border-radius: 12px;
            margin-top: 20px;
        }
    `;
    document.head.appendChild(style);
}

// ==========================================
// AUTHENTICATION CHECK
// ==========================================

// Check if user is logged in on page load
if (!localStorage.getItem('token')) {
    console.log('❌ No token found, redirecting to login');
    window.location.href = '/';
} else {
    console.log('✅ User is authenticated');
}

// Log that editor is loaded
console.log('✅ Editor.js loaded successfully');
console.log('📊 Initial state:', {
    theme: selectedTheme,
    imageStyle: selectedImageStyle,
    token: localStorage.getItem('token') ? 'Present' : 'Missing'
});