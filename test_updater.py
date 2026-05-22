#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script kiểm tra hệ thống Auto-Update

Chạy: python test_updater.py

Chức năng:
- Kiểm tra cấu hình GitHub
- Kiểm tra kết nối GitHub API
- Kiểm tra version management
- Kiểm tra update logic
"""

import sys
import json
from pathlib import Path

# Add app dir to path
sys.path.insert(0, str(Path(__file__).parent))

from updater import (
    AutoUpdater, UpdaterConfig, VersionManager, 
    GitHubChecker, Downloader, Installer
)


def print_header(text):
    """Print section header"""
    print("\n" + "="*60)
    print(f"  {text}")
    print("="*60)


def print_test(name, passed, error=None):
    """Print test result"""
    status = "✅ PASS" if passed else "❌ FAIL"
    print(f"{status}: {name}")
    if error:
        print(f"     Error: {error}")


def test_config():
    """Test configuration"""
    print_header("Testing Configuration")
    
    config = UpdaterConfig()
    
    # Test 1: Config file exists
    config_file = Path('github_config.json')
    test_passed = config_file.exists()
    print_test("Config file exists", test_passed)
    
    # Test 2: Load config
    try:
        owner = config.get('github_owner')
        repo = config.get('github_repo')
        is_valid = owner and repo and owner != "YOUR_GITHUB_USERNAME"
        print_test("Config values valid", is_valid, 
                  f"owner={owner}, repo={repo}")
        return is_valid
    except Exception as e:
        print_test("Config values valid", False, str(e))
        return False


def test_version_manager():
    """Test version management"""
    print_header("Testing Version Manager")
    
    # Test 1: Get version
    try:
        version = VersionManager.get_current_version()
        is_valid = version and '.' in version
        print_test("Get version", is_valid, f"version={version}")
    except Exception as e:
        print_test("Get version", False, str(e))
        return False
    
    # Test 2: Version format
    try:
        parts = version.split('.')
        is_valid = len(parts) >= 2 and all(p.isdigit() for p in parts)
        print_test("Version format valid (X.Y.Z)", is_valid, f"version={version}")
    except Exception as e:
        print_test("Version format valid", False, str(e))
        return False
    
    # Test 3: Compare versions
    try:
        result = VersionManager.compare_versions("1.0.0", "1.0.1")
        is_valid = result == -1  # 1.0.0 < 1.0.1
        print_test("Version comparison", is_valid)
    except Exception as e:
        print_test("Version comparison", False, str(e))
        return False
    
    return True


def test_github_checker():
    """Test GitHub checker"""
    print_header("Testing GitHub Checker")
    
    config = UpdaterConfig()
    owner = config.get('github_owner')
    repo = config.get('github_repo')
    
    # Check config
    if owner == "YOUR_GITHUB_USERNAME" or repo == "YOUR_REPO_NAME":
        print("⚠️  WARNING: GitHub config not properly set up!")
        print(f"   Current: owner={owner}, repo={repo}")
        print("   Please update github_config.json")
        return False
    
    # Test 1: Get latest release
    try:
        checker = GitHubChecker(config)
        release = checker.get_latest_release()
        
        if release is None:
            print_test("Get latest release", False, "No release found or API error")
            return False
        
        print_test("Get latest release", True)
        print(f"   Release: {release['tag_name']}")
        print(f"   Name: {release['name']}")
    except Exception as e:
        print_test("Get latest release", False, str(e))
        return False
    
    # Test 2: Check for updates
    try:
        has_update, info = checker.has_update()
        print_test("Check for updates", True, 
                  f"has_update={has_update}")
        if has_update:
            print(f"   New version available: {info['tag_name']}")
    except Exception as e:
        print_test("Check for updates", False, str(e))
        return False
    
    return True


def test_updater_config():
    """Test updater config"""
    print_header("Testing Updater Configuration")
    
    try:
        config = UpdaterConfig()
        
        # Check all settings
        settings = {
            'github_owner': config.get('github_owner'),
            'github_repo': config.get('github_repo'),
            'check_interval_hours': config.get('check_interval_hours'),
            'auto_download': config.get('auto_download'),
            'auto_install': config.get('auto_install'),
            'backup_before_update': config.get('backup_before_update'),
            'show_changelog': config.get('show_changelog')
        }
        
        print("Current settings:")
        for key, value in settings.items():
            print(f"  {key}: {value}")
        
        # Validate
        all_valid = all(settings.values() for key, value in settings.items() 
                       if key not in ['auto_download', 'auto_install', 'backup_before_update', 'show_changelog'])
        
        print_test("All settings configured", all_valid)
        return all_valid
    except Exception as e:
        print_test("Configuration", False, str(e))
        return False


def test_directories():
    """Test required directories"""
    print_header("Testing Directories")
    
    dirs = {
        'App directory': Path('.'),
        'Backups directory': Path('backups'),
    }
    
    all_exist = True
    for name, path in dirs.items():
        exists = path.exists()
        print_test(f"{name} exists", exists, str(path))
        if not exists:
            all_exist = False
    
    return all_exist


def test_files():
    """Test required files"""
    print_header("Testing Required Files")
    
    files = {
        'updater.py': Path('updater.py'),
        'launcher.py': Path('launcher.py'),
        'check_update.py': Path('check_update.py'),
        'version.txt': Path('version.txt'),
        'github_config.json': Path('github_config.json'),
    }
    
    all_exist = True
    for name, path in files.items():
        exists = path.exists()
        print_test(f"{name} exists", exists)
        if not exists:
            all_exist = False
    
    return all_exist


def test_imports():
    """Test Python imports"""
    print_header("Testing Python Imports")
    
    imports = {
        'urllib': 'urllib.request',
        'json': 'json',
        'pathlib': 'pathlib.Path',
        'threading': 'threading',
        'zipfile': 'zipfile',
        'tempfile': 'tempfile',
    }
    
    all_ok = True
    for name, module in imports.items():
        try:
            __import__(module.split('.')[0])
            print_test(f"Import {name}", True)
        except ImportError as e:
            print_test(f"Import {name}", False, str(e))
            all_ok = False
    
    return all_ok


def main():
    """Run all tests"""
    print("\n" + "🧪 AUTO-UPDATER SYSTEM TEST".center(60))
    
    tests = [
        ("Python Imports", test_imports),
        ("Required Files", test_files),
        ("Directories", test_directories),
        ("Version Manager", test_version_manager),
        ("Configuration", test_updater_config),
        ("GitHub Checker", test_github_checker),
        ("Custom Config", test_config),
    ]
    
    results = {}
    for test_name, test_func in tests:
        try:
            result = test_func()
            results[test_name] = result
        except Exception as e:
            print(f"❌ Unexpected error in {test_name}: {e}")
            results[test_name] = False
    
    # Summary
    print_header("Test Summary")
    
    passed = sum(1 for v in results.values() if v)
    total = len(results)
    
    for test_name, result in results.items():
        status = "✅" if result else "❌"
        print(f"{status} {test_name}")
    
    print(f"\nTotal: {passed}/{total} tests passed")
    
    if passed == total:
        print("\n✨ All tests passed! System is ready.")
        print("Next step: python launcher.py")
        return 0
    else:
        print("\n⚠️  Some tests failed. Please fix the issues above.")
        return 1


if __name__ == '__main__':
    exit_code = main()
    sys.exit(exit_code)
