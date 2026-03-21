/**
 * Authentication Logic
 * Gamma AI - Python Flask Edition
 */

// Check if user is already logged in
function checkAuth() {
    const token = localStorage.getItem('token');
    const currentPath = window.location.pathname;
    
    // If no token and not on login page, redirect to login
    if (!token && currentPath !== '/' && currentPath !== '/index.html') {
        window.location.href = '/';
        return false;
    }
    
    // If has token and on login page, redirect to dashboard
    if (token && (currentPath === '/' || currentPath === '/index.html')) {
        window.location.href = '/dashboard';
        return false;
    }
    
    return true;
}

/**
 * Handle login form submission
 */
document.getElementById('loginForm')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const email = document.getElementById('email').value;
    const password = document.getElementById('password').value;
    const loginBtn = document.getElementById('loginBtn');
    
    // Disable button
    loginBtn.disabled = true;
    loginBtn.textContent = '⏳ Logging in...';
    
    try {
        const data = await API.login(email, password);
        
        // Store token and user data
        localStorage.setItem('token', data.token);
        localStorage.setItem('user', JSON.stringify(data.user));
        
        // Show success message
        showToast('✅ Login successful! Redirecting...', 'success');
        
        // Redirect to dashboard after short delay
        setTimeout(() => {
            window.location.href = '/dashboard';
        }, 1000);
        
    } catch (error) {
        showToast('❌ ' + error.message, 'error');
        loginBtn.disabled = false;
        loginBtn.textContent = '🚀 Login to Gamma AI';
    }
});

/**
 * Handle signup form submission
 */
document.getElementById('signupForm')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const name = document.getElementById('name').value;
    const email = document.getElementById('email').value;
    const password = document.getElementById('password').value;
    const confirmPassword = document.getElementById('confirmPassword')?.value;
    const signupBtn = document.getElementById('signupBtn');
    
    // Validate passwords match
    if (confirmPassword && password !== confirmPassword) {
        showToast('❌ Passwords do not match!', 'error');
        return;
    }
    
    // Disable button
    signupBtn.disabled = true;
    signupBtn.textContent = '⏳ Creating account...';
    
    try {
        const data = await API.signup(name, email, password);
        
        // Store token and user data
        localStorage.setItem('token', data.token);
        localStorage.setItem('user', JSON.stringify(data.user));
        
        // Show success message
        showToast('✅ Account created! Redirecting...', 'success');
        
        // Redirect to dashboard
        setTimeout(() => {
            window.location.href = '/dashboard';
        }, 1000);
        
    } catch (error) {
        showToast('❌ ' + error.message, 'error');
        signupBtn.disabled = false;
        signupBtn.textContent = '🚀 Create Account';
    }
});

/**
 * Show signup form
 */
function showSignup() {
    const html = `
        <div class="modal-overlay" id="signupModal" onclick="closeSignup(event)">
            <div class="modal-content" onclick="event.stopPropagation()">
                <div class="modal-header">
                    <h2>Create Account</h2>
                    <button class="close-btn" onclick="closeSignup()">✕</button>
                </div>
                <form id="signupForm">
                    <div class="form-group">
                        <label for="name">Full Name</label>
                        <input type="text" id="name" required>
                    </div>
                    <div class="form-group">
                        <label for="email">Email</label>
                        <input type="email" id="email" required>
                    </div>
                    <div class="form-group">
                        <label for="password">Password</label>
                        <input type="password" id="password" required>
                    </div>
                    <div class="form-group">
                        <label for="confirmPassword">Confirm Password</label>
                        <input type="password" id="confirmPassword" required>
                    </div>
                    <button type="submit" class="btn btn-primary" id="signupBtn">
                        Create Account
                    </button>
                </form>
            </div>
        </div>
    `;
    
    document.body.insertAdjacentHTML('beforeend', html);
    
    // Add event listener to form
    document.getElementById('signupForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        
        const name = document.getElementById('name').value;
        const email = document.getElementById('email').value;
        const password = document.getElementById('password').value;
        const confirmPassword = document.getElementById('confirmPassword').value;
        const signupBtn = document.getElementById('signupBtn');
        
        if (password !== confirmPassword) {
            showToast('❌ Passwords do not match!', 'error');
            return;
        }
        
        signupBtn.disabled = true;
        signupBtn.textContent = '⏳ Creating account...';
        
        try {
            const data = await API.signup(name, email, password);
            
            localStorage.setItem('token', data.token);
            localStorage.setItem('user', JSON.stringify(data.user));
            
            showToast('✅ Account created successfully!', 'success');
            
            setTimeout(() => {
                window.location.href = '/dashboard';
            }, 1000);
            
        } catch (error) {
            showToast('❌ ' + error.message, 'error');
            signupBtn.disabled = false;
            signupBtn.textContent = 'Create Account';
        }
    });
}

/**
 * Close signup modal
 */
function closeSignup(event) {
    if (!event || event.target.id === 'signupModal') {
        const modal = document.getElementById('signupModal');
        if (modal) {
            modal.remove();
        }
    }
}

/**
 * Logout user
 */
function logout() {
    if (confirm('Are you sure you want to logout?')) {
        localStorage.removeItem('token');
        localStorage.removeItem('user');
        showToast('👋 Logged out successfully', 'success');
        setTimeout(() => {
            window.location.href = '/';
        }, 500);
    }
}

// Run auth check on page load
checkAuth();

// Add modal styles
if (!document.getElementById('modal-styles')) {
    const style = document.createElement('style');
    style.id = 'modal-styles';
    style.textContent = `
        .modal-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0, 0, 0, 0.7);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 9999;
            animation: fadeIn 0.3s ease;
        }
        
        .modal-content {
            background: white;
            border-radius: 20px;
            padding: 40px;
            max-width: 500px;
            width: 90%;
            max-height: 90vh;
            overflow-y: auto;
            animation: slideUp 0.3s ease;
        }
        
        .modal-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 30px;
        }
        
        .modal-header h2 {
            margin: 0;
            font-size: 28px;
            color: #1e293b;
        }
        
        .close-btn {
            background: none;
            border: none;
            font-size: 28px;
            cursor: pointer;
            color: #64748b;
            padding: 0;
            width: 32px;
            height: 32px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 8px;
            transition: all 0.3s;
        }
        
        .close-btn:hover {
            background: #f1f5f9;
            color: #1e293b;
        }
        
        @keyframes fadeIn {
            from { opacity: 0; }
            to { opacity: 1; }
        }
        
        @keyframes slideUp {
            from {
                transform: translateY(50px);
                opacity: 0;
            }
            to {
                transform: translateY(0);
                opacity: 1;
            }
        }
    `;
    document.head.appendChild(style);
}