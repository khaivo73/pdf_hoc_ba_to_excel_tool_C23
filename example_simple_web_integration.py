#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ví dụ đơn giản: Tích hợp Auto-Update vào Web App

Chạy: python example_simple_web_integration.py

Truy cập: http://localhost:5000

Tính năng:
- Web app với Flask
- Endpoint /api/version - lấy version
- Endpoint /api/check-update - kiểm tra update
- Endpoint /api/trigger-update - bắt đầu update
- Frontend hiển thị notification nếu có update mới
"""

from flask import Flask, render_template, jsonify
import threading
import sys
from pathlib import Path

# Import updater
sys.path.insert(0, str(Path(__file__).parent))
from updater import AutoUpdater, VersionManager

app = Flask(__name__)

# Global status
update_status = {
    'is_updating': False,
    'progress': 0,
    'message': ''
}


# ============ Routes ============

@app.route('/')
def index():
    """Main page"""
    return render_template_string('''
    <!DOCTYPE html>
    <html>
    <head>
        <title>HocBaPdfToExcel - Web Version</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                max-width: 800px;
                margin: 50px auto;
                padding: 20px;
                background: #f5f5f5;
            }
            .container {
                background: white;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }
            h1 { color: #333; }
            .version-info {
                padding: 10px;
                background: #e7f3ff;
                border-left: 4px solid #2196F3;
                margin: 20px 0;
                border-radius: 4px;
            }
            .update-notification {
                padding: 15px;
                background: #fff3cd;
                border: 1px solid #ffc107;
                border-radius: 4px;
                margin: 20px 0;
                display: none;
            }
            .update-notification.show {
                display: block;
            }
            .update-button {
                background: #007bff;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 4px;
                cursor: pointer;
                font-size: 14px;
                margin-top: 10px;
            }
            .update-button:hover {
                background: #0056b3;
            }
            .loading {
                display: none;
                text-align: center;
                padding: 20px;
            }
            .spinner {
                border: 4px solid #f3f3f3;
                border-top: 4px solid #2196F3;
                border-radius: 50%;
                width: 40px;
                height: 40px;
                animation: spin 1s linear infinite;
                margin: 0 auto;
            }
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>🔄 HocBaPdfToExcel - Web Version</h1>
            
            <div class="version-info" id="version-info">
                <strong>Version:</strong> <span id="current-version">Loading...</span>
            </div>
            
            <div class="update-notification" id="update-notification">
                <strong>📦 New Version Available!</strong>
                <p>
                    Current: <span id="current-ver"></span> →
                    New: <span id="new-ver"></span>
                </p>
                <p id="changelog" style="font-size: 12px; color: #666;"></p>
                <button class="update-button" onclick="triggerUpdate()">Update Now</button>
            </div>
            
            <div class="loading" id="loading">
                <div class="spinner"></div>
                <p id="loading-message">Updating...</p>
            </div>
            
            <h2>Features</h2>
            <ul>
                <li>Convert PDF to Excel</li>
                <li>Automatic updates from GitHub</li>
                <li>Web-based interface</li>
            </ul>
            
            <h2>API Endpoints</h2>
            <ul>
                <li><code>GET /api/version</code> - Get current version</li>
                <li><code>GET /api/check-update</code> - Check for updates</li>
                <li><code>POST /api/trigger-update</code> - Start update</li>
                <li><code>GET /api/update-status</code> - Get update status</li>
            </ul>
        </div>
        
        <script>
            // Check version on load
            document.addEventListener('DOMContentLoaded', function() {
                checkVersion();
                checkForUpdates();
                setInterval(checkForUpdates, 3600000); // Check every hour
            });
            
            async function checkVersion() {
                try {
                    const response = await fetch('/api/version');
                    const data = await response.json();
                    document.getElementById('current-version').textContent = data.version;
                } catch (error) {
                    console.error('Error:', error);
                }
            }
            
            async function checkForUpdates() {
                try {
                    const response = await fetch('/api/check-update');
                    const data = await response.json();
                    
                    if (data.has_update) {
                        document.getElementById('current-ver').textContent = data.current_version;
                        document.getElementById('new-ver').textContent = data.new_version;
                        document.getElementById('changelog').textContent = data.changelog;
                        document.getElementById('update-notification').classList.add('show');
                    }
                } catch (error) {
                    console.error('Error checking updates:', error);
                }
            }
            
            async function triggerUpdate() {
                if (!confirm('Update application? This may take a few minutes.')) {
                    return;
                }
                
                try {
                    document.getElementById('loading').style.display = 'block';
                    document.getElementById('loading-message').textContent = 'Downloading update...';
                    
                    const response = await fetch('/api/trigger-update', {method: 'POST'});
                    const data = await response.json();
                    
                    // Poll status
                    let checking = true;
                    while (checking) {
                        await new Promise(r => setTimeout(r, 2000));
                        
                        const statusResponse = await fetch('/api/update-status');
                        const statusData = await statusResponse.json();
                        
                        document.getElementById('loading-message').textContent = 
                            statusData.message || 'Updating...';
                        
                        if (!statusData.is_updating) {
                            checking = false;
                            if (statusData.success) {
                                alert('Update completed! Refreshing page...');
                                location.reload();
                            } else {
                                alert('Update failed: ' + statusData.error);
                            }
                        }
                    }
                } catch (error) {
                    alert('Error: ' + error);
                    document.getElementById('loading').style.display = 'none';
                }
            }
        </script>
    </body>
    </html>
    ''')


def render_template_string(html):
    """Render HTML string"""
    from flask import Flask
    return html


@app.route('/api/version', methods=['GET'])
def api_version():
    """Get current version"""
    return jsonify({
        'version': VersionManager.get_current_version()
    })


@app.route('/api/check-update', methods=['GET'])
def api_check_update():
    """Check for updates"""
    try:
        updater = AutoUpdater()
        has_update, release_info = updater.checker.has_update()
        
        if has_update:
            return jsonify({
                'has_update': True,
                'current_version': VersionManager.get_current_version(),
                'new_version': release_info['tag_name'],
                'release_name': release_info['name'],
                'changelog': release_info['body'][:300] + '...' if release_info['body'] else 'No changelog'
            })
        else:
            return jsonify({
                'has_update': False,
                'current_version': VersionManager.get_current_version()
            })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/trigger-update', methods=['POST'])
def api_trigger_update():
    """Start update (background)"""
    def run_update():
        global update_status
        try:
            update_status['is_updating'] = True
            update_status['message'] = 'Starting update...'
            update_status['success'] = False
            
            updater = AutoUpdater()
            
            def callback(status, message):
                global update_status
                update_status['message'] = message
                if status == 'update_complete':
                    update_status['success'] = True
                    update_status['is_updating'] = False
            
            updater.check_and_update(show_dialog=False, callback=callback)
            
        except Exception as e:
            update_status['error'] = str(e)
            update_status['is_updating'] = False
    
    # Start in background thread
    thread = threading.Thread(target=run_update, daemon=True)
    thread.start()
    
    return jsonify({'status': 'update_started'})


@app.route('/api/update-status', methods=['GET'])
def api_update_status():
    """Get update status"""
    return jsonify(update_status)


# ============ Main ============

if __name__ == '__main__':
    print("\n" + "="*60)
    print("🌐 HocBaPdfToExcel Web App")
    print("="*60)
    print(f"📦 Version: {VersionManager.get_current_version()}")
    print("🚀 Starting server on http://localhost:5000")
    print("\nPress Ctrl+C to stop")
    print("="*60 + "\n")
    
    app.run(debug=True, port=5000)
