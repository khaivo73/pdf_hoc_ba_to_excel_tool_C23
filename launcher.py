# -*- coding: utf-8 -*-
"""
Launcher cho HocBaPdfToExcel
- Kiểm tra và cập nhật tự động trước khi chạy ứng dụng
- Khởi động ứng dụng chính (app.py)

Sử dụng:
    python launcher.py [--skip-update] [--web]
    
    --skip-update: Bỏ qua kiểm tra cập nhật
    --web: Chạy phiên bản web (web_app.py)
"""

import sys
import os
import subprocess
from pathlib import Path
from updater import AutoUpdater, VersionManager, UpdaterConfig

APP_DIR = Path(__file__).resolve().parent


def main():
    """Hàm chính"""
    
    # Parse arguments
    skip_update = '--skip-update' in sys.argv
    run_web = '--web' in sys.argv
    
    # Hiển thị thông tin
    print("\n" + "="*60)
    print("🚀 HocBaPdfToExcel Launcher")
    print("="*60)
    print(f"📍 Thư mục ứng dụng: {APP_DIR}")
    print(f"📦 Version hiện tại: {VersionManager.get_current_version()}")
    
    # Kiểm tra cấu hình
    config = UpdaterConfig()
    owner = config.get("github_owner")
    repo = config.get("github_repo")
    
    if owner == "YOUR_GITHUB_USERNAME" or repo == "YOUR_REPO_NAME":
        print("\n⚠️  CẢNH BÁO: Cấu hình GitHub chưa được thiết lập!")
        print("📝 Hãy chỉnh sửa file 'github_config.json' với thông tin GitHub của bạn")
    
    # Kiểm tra cập nhật
    if not skip_update:
        print("\n🔄 Kiểm tra cập nhật...")
        try:
            updater = AutoUpdater()
            if not updater.check_and_update(show_dialog=True):
                print("⚠️  Kiểm tra cập nhật không thành công, tiếp tục khởi động...")
        except Exception as e:
            print(f"❌ Lỗi kiểm tra cập nhật: {e}")
            print("⏭️  Tiếp tục khởi động...")
    
    # Khởi động ứng dụng chính
    print("\n" + "="*60)
    print("🎬 Khởi động ứng dụng...")
    print("="*60 + "\n")
    
    try:
        if run_web:
            # Chạy web app
            app_file = APP_DIR / "web_app.py"
            subprocess.run([sys.executable, str(app_file)], cwd=APP_DIR, check=True)
        else:
            # Chạy desktop app
            app_file = APP_DIR / "app.py"
            subprocess.run([sys.executable, str(app_file)], cwd=APP_DIR, check=True)
    
    except FileNotFoundError:
        print("❌ Lỗi: Không tìm thấy file ứng dụng")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n👋 Đã dừng ứng dụng")
        sys.exit(0)
    except Exception as e:
        print(f"❌ Lỗi khi chạy ứng dụng: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
