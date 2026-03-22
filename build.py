#!/usr/bin/env python3
"""
Build script for Katydid Analyzer executables.
- On Mac: creates .app bundles (Wav Analyzer.app, Data Analyzer.app)
- On Windows: creates .exe files (Wav Analyzer.exe, Data Analyzer.exe)

Run: python build.py
Output: dist/ folder in project directory
"""

import subprocess
import sys
import os

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    # Ensure PyInstaller is available
    try:
        import PyInstaller
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        import PyInstaller

    # Build both apps
    specs = [
        ("wav_analyzer.spec", "Wav Analyzer"),
        ("data_analyzer.spec", "Data Analyzer"),
    ]

    for spec_file, app_name in specs:
        print(f"\n{'='*50}")
        print(f"Building {app_name}...")
        print(f"{'='*50}")
        result = subprocess.run(
            [sys.executable, "-m", "PyInstaller", "--clean", "--noconfirm", spec_file],
            cwd=script_dir,
        )
        if result.returncode != 0:
            print(f"ERROR: Build failed for {app_name}")
            sys.exit(1)

    # Copy outputs to a release folder for easy sharing
    dist_dir = os.path.join(script_dir, "dist")
    release_dir = os.path.join(script_dir, "release")
    os.makedirs(release_dir, exist_ok=True)

    if sys.platform == "darwin":
        # Mac: copy .app bundles
        for _, app_name in specs:
            src = os.path.join(dist_dir, f"{app_name}.app")
            dst = os.path.join(release_dir, f"{app_name}.app")
            if os.path.exists(src):
                if os.path.exists(dst):
                    import shutil
                    shutil.rmtree(dst)
                import shutil
                shutil.copytree(src, dst)
                print(f"  -> {release_dir}/{app_name}.app")
    else:
        # Windows: copy .exe files
        for _, app_name in specs:
            src = os.path.join(dist_dir, f"{app_name}.exe")
            dst = os.path.join(release_dir, f"{app_name}.exe")
            if os.path.exists(src):
                import shutil
                shutil.copy2(src, dst)
                print(f"  -> {release_dir}/{app_name}.exe")

    print(f"\nDone! Executables are in: {release_dir}")

if __name__ == "__main__":
    main()
