"""
Automated PyInstaller Build Script for SLAM Jobcard Auto-Filler
Produces a standalone Windows executable (.exe) that works on any Windows PC.
"""
import os
import sys
import shutil
import subprocess
import zipfile

def build():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    dist_dir = os.path.join(base_dir, "dist")
    build_dir = os.path.join(base_dir, "build")
    spec_file = os.path.join(base_dir, "SLAM_Auto_Filler.spec")

    print("=" * 60)
    print("SLAM Jobcard Auto-Filler - Executable Builder")
    print("=" * 60)

    # Clean old build artifacts if present
    for p in [build_dir, spec_file]:
        if os.path.isdir(p):
            try:
                shutil.rmtree(p)
            except Exception as e:
                print(f"Warning: Could not remove {p}: {e}")
        elif os.path.isfile(p):
            try:
                os.remove(p)
            except Exception as e:
                print(f"Warning: Could not remove {p}: {e}")

    entry_point = os.path.join(base_dir, "main.py")

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        entry_point,
        "--name=SLAM_Auto_Filler",
        "--onefile",
        "--windowed",
        "--clean",
        "--noconfirm",
        "--collect-all=playwright",
        "--hidden-import=openpyxl",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.ttk",
        "--hidden-import=tkinter.messagebox",
        "--hidden-import=tkinter.filedialog",
        "--hidden-import=excel_parser",
        "--hidden-import=slam_bot",
        "--hidden-import=gui",
        "--hidden-import=config",
        "--hidden-import=updater",
    ]

    print(f"\nRunning PyInstaller command:\n{' '.join(cmd)}\n")
    ret = subprocess.run(cmd, cwd=base_dir)

    if ret.returncode != 0:
        print("\n[ERROR] Build failed! Check PyInstaller errors above.")
        sys.exit(ret.returncode)

    exe_path = os.path.join(dist_dir, "SLAM_Auto_Filler.exe")
    if os.path.exists(exe_path):
        root_exe = os.path.join(base_dir, "SLAM_Auto_Filler.exe")
        try:
            shutil.copy2(exe_path, root_exe)
            print(f"Copied to project root: {root_exe}")
        except Exception as e:
            print(f"Warning: Could not copy to root: {e}")

        size_mb = os.path.getsize(exe_path) / (1024 * 1024)
        print("\n" + "=" * 60)
        print("BUILD SUCCESSFUL!")
        print(f"Executable: {exe_path}")
        print(f"Size: {size_mb:.2f} MB")
        print("=" * 60)
    else:
        print(f"\n[WARNING] {exe_path} was not found.")

if __name__ == "__main__":
    build()
