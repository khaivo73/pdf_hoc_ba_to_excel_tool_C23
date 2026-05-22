# -*- coding: utf-8 -*-
"""
Auto updater cho ứng dụng HocBaPdfToExcel
- Kiểm tra version mới từ GitHub Releases
- Tải và cài đặt tự động
- Hỗ trợ cả phiên bản Python và EXE

Tác giả: Auto-Updater System
"""

import os
import sys
import json
import shutil
import subprocess
import threading
import urllib.request
import urllib.error
from pathlib import Path
from typing import Optional, Dict, Tuple
import zipfile
import tempfile
from datetime import datetime

# Thêm thư mục hiện tại vào path
APP_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(APP_DIR))


class UpdaterConfig:
    """Cấu hình cho updater"""
    
    def __init__(self, config_file: Optional[Path] = None):
        if config_file is None:
            config_file = APP_DIR / "github_config.json"
        
        self.config_file = config_file
        self.config = self._load_config()
        
    def _load_config(self) -> Dict:
        """Tải cấu hình từ file JSON"""
        if self.config_file.exists():
            with open(self.config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        return self._default_config()
    
    def _default_config(self) -> Dict:
        """Cấu hình mặc định"""
        return {
            "github_owner": "YOUR_GITHUB_USERNAME",
            "github_repo": "YOUR_REPO_NAME",
            "check_interval_hours": 24,
            "auto_download": True,
            "auto_install": False,  # Thay vì tự động cài, hỏi người dùng
            "backup_before_update": True,
            "show_changelog": True
        }
    
    def get(self, key: str, default=None):
        """Lấy giá trị config"""
        return self.config.get(key, default)
    
    @property
    def github_api_url(self) -> str:
        """URL API GitHub để lấy releases"""
        owner = self.get("github_owner")
        repo = self.get("github_repo")
        return f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
    
    @property
    def github_releases_url(self) -> str:
        """URL trang releases GitHub"""
        owner = self.get("github_owner")
        repo = self.get("github_repo")
        return f"https://github.com/{owner}/{repo}/releases"


class VersionManager:
    """Quản lý version"""
    
    VERSION_FILE = APP_DIR / "version.txt"
    
    @staticmethod
    def get_current_version() -> str:
        """Lấy version hiện tại"""
        if VersionManager.VERSION_FILE.exists():
            with open(VersionManager.VERSION_FILE, 'r', encoding='utf-8') as f:
                return f.read().strip()
        return "0.0.0"
    
    @staticmethod
    def set_version(version: str):
        """Lưu version mới"""
        with open(VersionManager.VERSION_FILE, 'w', encoding='utf-8') as f:
            f.write(version)
    
    @staticmethod
    def compare_versions(v1: str, v2: str) -> int:
        """
        So sánh 2 version
        Return:
            -1: v1 < v2 (có version mới)
            0: v1 == v2
            1: v1 > v2
        """
        try:
            parts1 = [int(x) for x in v1.split('.')]
            parts2 = [int(x) for x in v2.split('.')]
            
            # Pad với 0 nếu cần
            max_len = max(len(parts1), len(parts2))
            parts1 += [0] * (max_len - len(parts1))
            parts2 += [0] * (max_len - len(parts2))
            
            for p1, p2 in zip(parts1, parts2):
                if p1 < p2:
                    return -1
                elif p1 > p2:
                    return 1
            return 0
        except:
            return 0


class GitHubChecker:
    """Kiểm tra update từ GitHub"""
    
    def __init__(self, config: UpdaterConfig):
        self.config = config
        self.timeout = 10  # giây
    
    def get_latest_release(self) -> Optional[Dict]:
        """
        Lấy thông tin release mới nhất từ GitHub
        Returns:
            Dict với: tag_name, name, body, download_url
            None nếu lỗi
        """
        try:
            url = self.config.github_api_url
            headers = {
                'User-Agent': 'HocBaPdfToExcel-Updater/1.0',
                'Accept': 'application/vnd.github.v3+json'
            }
            
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=self.timeout) as response:
                data = json.loads(response.read().decode('utf-8'))
            
            # Tìm download URL cho file .zip hoặc .exe
            download_url = self._find_download_url(data.get('assets', []))
            
            return {
                'tag_name': data.get('tag_name', 'unknown'),
                'name': data.get('name', 'Release'),
                'body': data.get('body', ''),
                'download_url': download_url,
                'published_at': data.get('published_at', '')
            }
        except urllib.error.URLError as e:
            print(f"❌ Lỗi kết nối GitHub: {e}")
            return None
        except Exception as e:
            print(f"❌ Lỗi khi kiểm tra release: {e}")
            return None
    
    def _find_download_url(self, assets: list) -> Optional[str]:
        """Tìm URL download cho asset"""
        # Ưu tiên file .zip, nếu không có thì .exe
        for asset in assets:
            if asset['name'].endswith('.zip'):
                return asset['browser_download_url']
        
        for asset in assets:
            if asset['name'].endswith('.exe'):
                return asset['browser_download_url']
        
        return None
    
    def has_update(self) -> Tuple[bool, Optional[Dict]]:
        """
        Kiểm tra xem có update mới
        Returns:
            (has_update, release_info)
        """
        current_version = VersionManager.get_current_version()
        release = self.get_latest_release()
        
        if not release or not release['download_url']:
            return False, None
        
        new_version = release['tag_name'].lstrip('v')  # Loại bỏ 'v' prefix
        
        if VersionManager.compare_versions(current_version, new_version) < 0:
            return True, release
        
        return False, None


class Downloader:
    """Tải file từ GitHub"""
    
    @staticmethod
    def download_file(url: str, destination: Path, callback=None) -> bool:
        """
        Tải file từ URL
        Args:
            url: URL file
            destination: Đường dẫn lưu file
            callback: Hàm callback(downloaded, total) để theo dõi tiến độ
        Returns:
            True nếu thành công
        """
        try:
            def report_progress(block_num, block_size, total_size):
                downloaded = block_num * block_size
                if callback:
                    callback(min(downloaded, total_size), total_size)
            
            print(f"📥 Tải file từ: {url}")
            urllib.request.urlretrieve(url, destination, reporthook=report_progress)
            print(f"✅ Tải xong: {destination}")
            return True
        except Exception as e:
            print(f"❌ Lỗi khi tải file: {e}")
            return False


class Installer:
    """Cài đặt update"""
    
    def __init__(self, config: UpdaterConfig):
        self.config = config
        self.backup_dir = APP_DIR / "backups"
        self.backup_dir.mkdir(exist_ok=True)
    
    def install_update(self, package_path: Path, version: str) -> bool:
        """
        Cài đặt update
        Args:
            package_path: Đường dẫn file .zip hoặc .exe
            version: Version mới
        Returns:
            True nếu thành công
        """
        try:
            # Backup trước nếu cấu hình cho phép
            if self.config.get("backup_before_update", True):
                if not self._backup_current():
                    print("⚠️  Không thể backup, tiếp tục...")
            
            # Kiểm tra loại file
            if package_path.suffix == '.zip':
                success = self._install_from_zip(package_path)
            elif package_path.suffix == '.exe':
                success = self._install_from_exe(package_path)
            else:
                print(f"❌ Loại file không hỗ trợ: {package_path.suffix}")
                return False
            
            if success:
                # Cập nhật version
                VersionManager.set_version(version)
                print(f"✅ Cập nhật lên version {version} thành công")
                return True
            
            return False
        except Exception as e:
            print(f"❌ Lỗi khi cài đặt: {e}")
            return False
    
    def _backup_current(self) -> bool:
        """Backup version hiện tại"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            current_version = VersionManager.get_current_version()
            backup_name = f"backup_{current_version}_{timestamp}"
            backup_path = self.backup_dir / backup_name
            
            backup_path.mkdir(exist_ok=True)
            
            # Backup các file chính
            important_files = ['app.py', 'web_app.py', 'requirements.txt']
            for fname in important_files:
                fpath = APP_DIR / fname
                if fpath.exists():
                    shutil.copy2(fpath, backup_path / fname)
            
            # Backup thư mục templates
            templates_dir = APP_DIR / 'templates'
            if templates_dir.exists():
                shutil.copytree(templates_dir, backup_path / 'templates', dirs_exist_ok=True)
            
            print(f"✅ Backup tại: {backup_path}")
            return True
        except Exception as e:
            print(f"⚠️  Lỗi khi backup: {e}")
            return False
    
    def _install_from_zip(self, zip_path: Path) -> bool:
        """Cài đặt từ file .zip"""
        try:
            print(f"📦 Giải nén: {zip_path}")
            
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)
                
                # Giải nén
                with zipfile.ZipFile(zip_path, 'r') as zf:
                    zf.extractall(temp_path)
                
                # Copy các file Python chính
                for item in temp_path.rglob('*.py'):
                    rel_path = item.relative_to(temp_path)
                    dest = APP_DIR / rel_path
                    dest.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(item, dest)
                
                # Copy requirements.txt nếu có
                req_file = None
                for item in temp_path.rglob('requirements.txt'):
                    req_file = item
                    break
                
                if req_file:
                    shutil.copy2(req_file, APP_DIR / 'requirements.txt')
                    # Cài đặt dependencies
                    self._install_dependencies()
            
            return True
        except Exception as e:
            print(f"❌ Lỗi khi giải nén: {e}")
            return False
    
    def _install_from_exe(self, exe_path: Path) -> bool:
        """Cài đặt từ file .exe"""
        try:
            print(f"📥 Chạy cài đặt: {exe_path}")
            subprocess.run([str(exe_path)], check=False)
            return True
        except Exception as e:
            print(f"❌ Lỗi khi chạy cài đặt: {e}")
            return False
    
    def _install_dependencies(self):
        """Cài đặt dependencies từ requirements.txt"""
        try:
            print("📦 Cài đặt dependencies...")
            subprocess.run(
                [sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt'],
                cwd=APP_DIR,
                check=False
            )
        except Exception as e:
            print(f"⚠️  Lỗi khi cài dependencies: {e}")


class AutoUpdater:
    """Lớp chính quản lý auto-update"""
    
    def __init__(self):
        self.config = UpdaterConfig()
        self.checker = GitHubChecker(self.config)
        self.installer = Installer(self.config)
        self.downloader = Downloader()
    
    def check_and_update(self, show_dialog=True, callback=None) -> bool:
        """
        Kiểm tra và cài đặt update nếu có
        Args:
            show_dialog: Hiển thị hộp thoại người dùng
            callback: Hàm callback(status, message)
        Returns:
            True nếu cập nhật thành công hoặc không cần cập nhật
        """
        print("\n" + "="*60)
        print("🔍 Kiểm tra update...")
        print("="*60)
        
        try:
            has_update, release_info = self.checker.has_update()
            
            if not has_update:
                msg = f"✅ Ứng dụng đã là phiên bản mới nhất ({VersionManager.get_current_version()})"
                print(msg)
                if callback:
                    callback('no_update', msg)
                return True
            
            # Có update mới
            current_ver = VersionManager.get_current_version()
            new_ver = release_info['tag_name'].lstrip('v')
            msg = f"\n📦 Có phiên bản mới: {new_ver} (hiện tại: {current_ver})"
            print(msg)
            print(f"📝 Thay đổi: {release_info['name']}")
            
            if self.config.get("show_changelog"):
                changelog = release_info.get('body', '')
                if changelog:
                    print(f"\n📋 Changelog:\n{changelog[:500]}...")
            
            if callback:
                callback('update_available', msg)
            
            # Hỏi người dùng (nếu show_dialog=True)
            if show_dialog:
                if not self._ask_user_confirmation(new_ver, release_info):
                    print("⏭️  Người dùng bỏ qua cập nhật")
                    return True
            
            # Tải file
            download_url = release_info['download_url']
            if not download_url:
                print("❌ Không tìm thấy file để tải")
                return False
            
            # Xác định tên file dựa trên URL
            temp_file = Path(tempfile.gettempdir()) / Path(download_url).name
            
            if not self.downloader.download_file(download_url, temp_file):
                return False
            
            if callback:
                callback('download_complete', f"Tải xong: {temp_file}")
            
            # Cài đặt
            if self.installer.install_update(temp_file, new_ver):
                print("\n✨ Cập nhật hoàn tất! Vui lòng khởi động lại ứng dụng.")
                if callback:
                    callback('update_complete', 'Cập nhật hoàn tất')
                
                # Xóa file tạm
                try:
                    temp_file.unlink()
                except:
                    pass
                
                return True
            
            return False
        
        except Exception as e:
            error_msg = f"❌ Lỗi: {e}"
            print(error_msg)
            if callback:
                callback('error', error_msg)
            return False
    
    def _ask_user_confirmation(self, new_version: str, release_info: Dict) -> bool:
        """Hỏi người dùng xác nhận cập nhật"""
        try:
            import tkinter as tk
            from tkinter import messagebox
            
            # Tạo cửa sổ ẩn để show messagebox
            root = tk.Tk()
            root.withdraw()
            
            msg = f"Có phiên bản mới {new_version}\n\n{release_info['name']}\n\nBạn muốn cập nhật ngay?"
            result = messagebox.askyesno("Cập nhật phần mềm", msg)
            root.destroy()
            
            return result
        except:
            # Nếu không có tkinter, hỏi qua console
            response = input("\nCập nhật ngay? (y/n): ").strip().lower()
            return response == 'y'
    
    def check_for_update_async(self, callback=None):
        """Kiểm tra update trong thread riêng"""
        def run():
            self.check_and_update(show_dialog=True, callback=callback)
        
        thread = threading.Thread(target=run, daemon=True)
        thread.start()


def setup_initial_config():
    """Tạo file cấu hình mẫu nếu chưa có"""
    config_file = APP_DIR / "github_config.json"
    
    if not config_file.exists():
        config_template = {
            "github_owner": "YOUR_GITHUB_USERNAME",
            "github_repo": "YOUR_REPO_NAME",
            "check_interval_hours": 24,
            "auto_download": True,
            "auto_install": False,
            "backup_before_update": True,
            "show_changelog": True
        }
        
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(config_template, f, indent=2, ensure_ascii=False)
        
        print(f"\n⚙️  Tạo file cấu hình: {config_file}")
        print("📝 Vui lòng chỉnh sửa file này với thông tin GitHub của bạn")


if __name__ == '__main__':
    # Test
    setup_initial_config()
    
    updater = AutoUpdater()
    updater.check_and_update(show_dialog=True)
