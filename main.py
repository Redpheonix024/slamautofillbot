"""
SLAM Jobcard Auto-Filler Bot - Entry Point
Supports both Desktop GUI and Command Line (CLI) execution.
"""
import sys
import argparse
import os

# Register Windows AppUserModelID before any UI / COM initialization
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("IndianRailways.SLAM.JobcardAutoFiller.v1")
    except Exception:
        pass

def main():
    parser = argparse.ArgumentParser(description="Indian Railways SLAM Jobcard Auto-Filler Bot")
    parser.add_argument("--file", "-f", default="", help="Path to bookings Excel file (.xlsx)")
    parser.add_argument("--loco", "-l", default="", help="Override locomotive number")
    parser.add_argument("--otp", default="", help="SLAM OTP (if available)")
    parser.add_argument("--type", default="", help="Jobcard Type (e.g. 'Schedule Attention')")
    parser.add_argument("--eqt", default="", help="Equipment Required (e.g. 'Jobcard Without Eq')")
    parser.add_argument("--source", default="", help="Source / Obs Type (e.g. 'Testing Remarks')")
    parser.add_argument("--kind", default="", help="Kind (e.g. 'Minor')")
    parser.add_argument("--headless", action="store_true", help="Run browser in background (hidden)")
    parser.add_argument("--cli", action="store_true", help="Force command-line mode")

    args = parser.parse_args()

    # If no file or --cli passed, launch desktop GUI
    if not args.file and not args.cli:
        try:
            from gui import main as run_gui
            run_gui()
        except Exception as e:
            try:
                import tkinter as tk
                from tkinter import messagebox
                r = tk.Tk()
                r.withdraw()
                messagebox.showerror("SLAM Auto-Filler Startup Error", f"A fatal error occurred while starting the application:\n\n{e}")
            except Exception:
                pass
            raise
        return

    # CLI Execution Mode
    file_path = args.file
    if not file_path:
        print("Error: Please provide --file 'path/to/file.xlsx'")
        sys.exit(1)

    if not os.path.exists(file_path):
        print(f"Error: File not found at '{file_path}'")
        sys.exit(1)

    import config
    from slam_bot import SLAMBot

    jc_type = args.type or config.DEFAULT_JOBCARD_TYPE
    eqt_req = args.eqt or config.DEFAULT_EQT_REQUIRED
    obs_type = args.source or config.DEFAULT_OBS_TYPE
    kind = args.kind or config.DEFAULT_JOB_KIND

    print("=" * 60)
    print("      INDIAN RAILWAYS - SLAM JOBCARD AUTO-FILLER (CLI)      ")
    print("=" * 60)
    print(f"File         : {file_path}")
    print(f"Jobcard Type : {jc_type}")
    print(f"Eqt Required : {eqt_req}")
    print(f"Source       : {obs_type}")
    print(f"Kind         : {kind}")
    print(f"Headless     : {args.headless}")
    print("=" * 60)

    bot = SLAMBot(
        jobcard_type=jc_type,
        eqt_required=eqt_req,
        obs_type=obs_type,
        job_kind=kind,
        headless=args.headless,
        verify_before_save=True
    )

    try:
        res = bot.process_excel(file_path, otp=args.otp)
        print("\n" + "=" * 60)
        print(f"STATUS  : {res.get('status', 'Unknown').upper()}")
        print(f"LOCO    : {res.get('loco_number')}")
        print(f"ENTERED : {res.get('entered')}/{res.get('total')}")
        if res.get("failed_items"):
            print(f"FAILED  : {len(res['failed_items'])}")
            for item in res["failed_items"]:
                print(f"   -> S.No {item['sno']}: {item['reason']}")
        print("=" * 60)
    except KeyboardInterrupt:
        print("\nInterrupted by user.")
    except Exception as e:
        print(f"\n[FATAL ERROR] {e}")
        sys.exit(1)
    finally:
        if not args.headless:
            try:
                input("\n[SLAM Bot] Browser left open for your review. Press Enter to close browser and exit...")
            except (KeyboardInterrupt, EOFError):
                pass
        bot.close()

if __name__ == "__main__":
    main()
