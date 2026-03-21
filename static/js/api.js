/**
 * GAMMA AI - API Helper Functions
 * Author: GuptaSigma | Date: 2025-11-23
 * Complete Enhanced Version with Toast Notifications
 */

const API_BASE = window.location.origin + '/api';

// ==========================================
// TOKEN & USER MANAGEMENT
// ==========================================
function getToken() {
    return localStorage.getItem('token');
}

function getUser() {
    const user = localStorage.getItem('user');
    return user ? JSON.parse(user) : null;
}

function isAuthenticated() {
    return !!getToken();
}

function setToken(token) {
    localStorage.setItem('token', token);
}

function setUser(user) {
    localStorage.setItem('user', JSON.stringify(user));
}

function clearAuth() {
    localStorage.removeItem('token');
    localStorage.removeItem('user');
}

// ==========================================
// API REQUEST HANDLER
// ==========================================
async function apiRequest(endpoint, options = {}) {
    const token = getToken();
    
    const defaultOptions = {
        headers: {
            'Content-Type': 'application/json',
            ...(token && { 'Authorization': `Bearer ${token}` })
        }
    };
    
    const mergedOptions = {
        ...defaultOptions,
        ...options,
        headers: {
            ...defaultOptions.headers,
            ...options.headers
        }
    };
    
    try {
        const response = await fetch(`${API_BASE}${endpoint}`, mergedOptions);
        const data = await response.json();
        
        if (!response.ok) {
            throw new Error(data.error || 'Request failed');
        }
        
        return { success: true, data };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

// ==========================================
// AUTH API
// ==========================================
const AuthAPI = {
    async signup(name, email, password) {
        return apiRequest('/auth/signup', {
            method: 'POST',
            body: JSON.stringify({ name, email, password })
        });
    },
    
    async login(email, password) {
        return apiRequest('/auth/login', {
            method: 'POST',
            body: JSON.stringify({ email, password })
        });
    },
    
    logout() {
        clearAuth();
        showToast('Logged out successfully', 'success');
        setTimeout(() => {
            window.location.href = '/';
        }, 1000);
    }
};

// ==========================================
// PRESENTATIONS API
// ==========================================
const PresentationsAPI = {
    async list() {
        return apiRequest('/presentations/');
    },
    
    async get(id) {
        return apiRequest(`/presentations/${id}`);
    },
    
    async generate(data) {
        return apiRequest('/presentations/generate', {
            method: 'POST',
            body: JSON.stringify(data)
        });
    },
    
    async update(id, data) {
        return apiRequest(`/presentations/${id}`, {
            method: 'PUT',
            body: JSON.stringify(data)
        });
    },
    
    async delete(id) {
        return apiRequest(`/presentations/${id}`, {
            method: 'DELETE'
        });
    },
    
    getExportUrl(id, format) {
        return `${API_BASE}/presentations/${id}/export?format=${format}`;
    }
};

// ==========================================
// TOAST NOTIFICATION FUNCTIONS
// ==========================================
function showToast(message, type = 'success') {
    // Check if toast exists, create if not
    let toast = document.getElementById('toast');
    
    if (!toast) {
        // Create toast element dynamically
        toast = document.createElement('div');
        toast.id = 'toast';
        toast.className = 'toast';
        toast.innerHTML = `
            <div class="toast-icon" id="toastIcon">✅</div>
            <div class="toast-message" id="toastMessage">Success!</div>
            <div class="toast-close" onclick="hideToast()">✕</div>
        `;
        document.body.appendChild(toast);
    }
    
    const icon = document.getElementById('toastIcon');
    const messageEl = document.getElementById('toastMessage');
    
    // Set content
    messageEl.textContent = message;
    
    // Remove all type classes
    toast.classList.remove('success', 'error', 'warning', 'info');
    
    // Set icon and style based on type
    switch(type) {
        case 'success':
            icon.textContent = '✅';
            toast.classList.add('success');
            break;
        case 'error':
            icon.textContent = '❌';
            toast.classList.add('error');
            break;
        case 'warning':
            icon.textContent = '⚠️';
            toast.classList.add('warning');
            break;
        case 'info':
            icon.textContent = 'ℹ️';
            toast.classList.add('info');
            break;
        default:
            icon.textContent = '✅';
            toast.classList.add('success');
    }
    
    // Show toast with animation
    toast.classList.add('show');
    
    // Auto hide after 3 seconds
    setTimeout(() => {
        hideToast();
    }, 3000);
}

function hideToast() {
    const toast = document.getElementById('toast');
    if (toast) {
        toast.classList.remove('show');
    }
}

// ==========================================
// LOADING OVERLAY FUNCTIONS
// ==========================================
function showLoadingOverlay() {
    let overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.classList.add('active');
    }
}

function hideLoadingOverlay() {
    let overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.classList.remove('active');
    }
}

// ==========================================
// BUTTON LOADER FUNCTIONS (Legacy Support)
// ==========================================
function showLoader(buttonElement) {
    const text = buttonElement.querySelector('.btn-text');
    const loader = buttonElement.querySelector('.btn-loader');
    
    if (text) text.style.display = 'none';
    if (loader) loader.style.display = 'inline';
    buttonElement.disabled = true;
}

function hideLoader(buttonElement) {
    const text = buttonElement.querySelector('.btn-text');
    const loader = buttonElement.querySelector('.btn-loader');
    
    if (text) text.style.display = 'inline';
    if (loader) loader.style.display = 'none';
    buttonElement.disabled = false;
}

// ==========================================
// UTILITY FUNCTIONS
// ==========================================
function formatDate(dateString) {
    const date = new Date(dateString);
    const options = { 
        year: 'numeric', 
        month: 'short', 
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
    };
    return date.toLocaleDateString('en-US', options);
}

function truncateText(text, maxLength = 100) {
    if (text.length <= maxLength) return text;
    return text.substring(0, maxLength) + '...';
}

function downloadFile(blob, filename) {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
}

// ==========================================
// INITIALIZATION
// ==========================================
console.log('✅ Gamma AI API Helper loaded');
console.log('📅 Date: 2025-11-23');
console.log('👤 Author: GuptaSigma');
console.log('🎨 Features: Toast Notifications, API Wrapper, Utilities');