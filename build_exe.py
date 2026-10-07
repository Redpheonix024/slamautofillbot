"""
Automated PyInstaller Build Script for SLAM Jobcard Auto-Filler
Produces a standalone Windows executable (.exe) that works on any Windows PC.
"""
import os
import sys
import shutil
import subprocess
import zipfile

def build(exe_base_name: str = "SLAM_Auto_Filler_v1.1_Equipment"):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    dist_dir = os.path.join(base_dir, "dist")
    build_dir = os.path.join(base_dir, "build_v1_1")
    spec_file = os.path.join(base_dir, f"{exe_base_name}.spec")

    print("=" * 60)
    print(f"SLAM Auto-Filler ({exe_base_name}) - Executable Builder")
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

    icon_file = os.path.join(base_dir, "app_icon.ico")
    png_file = os.path.join(base_dir, "app_icon.png")

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        entry_point,
        f"--name={exe_base_name}",
        f"--workpath={build_dir}",
        "--onefile",
        "--windowed",
        "--clean",
        "--noconfirm",
        f"--icon={icon_file}",
        f"--add-data={icon_file};.",
        f"--add-data={png_file};.",
        "--collect-all=playwright",
        "--hidden-import=PIL",
        "--hidden-import=PIL.ImageTk",
        "--hidden-import=PIL.Image",
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

    exe_path = os.path.join(dist_dir, f"{exe_base_name}.exe")
    if os.path.exists(exe_path):
        root_exe = os.path.join(base_dir, f"{exe_base_name}.exe")
        try:
            shutil.copy2(exe_path, root_exe)
            print(f"Copied to project root: {root_exe}")
        except Exception as e:
            print(f"Warning: Could not copy to root: {e}")

        size_mb = os.path.getsize(exe_path) / (1024 * 1024)
        print("\n" + "=" * 60)
        print("BUILD SUCCESSFUL!")
        print(f"Branched Executable: {root_exe}")
        print(f"Original SLAM_Auto_Filler.exe was left UNTOUCHED.")
        print(f"Size: {size_mb:.2f} MB")
        print("=" * 60)
    else:
        print(f"\n[WARNING] {exe_path} was not found.")

if __name__ == "__main__":
    build()
