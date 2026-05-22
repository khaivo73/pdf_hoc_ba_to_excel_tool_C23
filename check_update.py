# -*- coding: utf-8 -*-
"""
Script riêng để kiểm tra update
Có thể được chạy định kỳ (via task scheduler, cron, etc.)

Sử dụng:
    python check_update.py
    python check_update.py --auto-install
"""

import sys
from pathlib import Path
from updater import AutoUpdater, UpdaterConfig, VersionManager

APP_DIR = Path(__file__).resolve().parent


def main():
    """Chạy kiểm tra update"""
    
    auto_install = '--auto-install' in sys.argv
    
    print("\n" + "="*60)
    print("🔍 Kiểm tra cập nhật HocBaPdfToExcel")
    print("="*60)
    print(f"📦 Version hiện tại: {VersionManager.get_current_version()}\n")
    
    # Kiểm tra cấu hình
    config = UpdaterConfig()
    owner = config.get("github_owner")
    repo = config.get("github_repo")
    
    if owner == "YOUR_GITHUB_USERNAME" or repo == "YOUR_REPO_NAME":
        print("❌ Lỗi: Cấu hình GitHub chưa được thiết lập!")
        print("📝 Vui lòng chỉnh sửa file 'github_config.json'")
        return False
    
    try:
        updater = AutoUpdater()
        updater.check_and_update(
            show_dialog=not auto_install,
            callback=_callback
        )
        return True
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        return False


def _callback(status: str, message: str):
    """Callback để xử lý các event"""
    if status == 'no_update':
        print(f"✅ {message}")
    elif status == 'update_available':
        print(f"📦 {message}")
    elif status == 'download_complete':
        print(f"✅ {message}")
    elif status == 'update_complete':
        print(f"✨ {message}")
    elif status == 'error':
        print(f"❌ {message}")


if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
