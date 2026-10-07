"""
Meeting Import GUI v1.5.3 - Modular Entry Point
Imports the main GUI class while keeping all functionality intact
"""

import sys


def main():
    """Main entry point."""
    
    # Step 1: Check and install dependencies BEFORE importing GUI
    print("Checking dependencies...")
    try:
        from dependency_manager import check_and_install_dependencies
        
        success, messages = check_and_install_dependencies()
        
        # Print all messages
        for msg in messages:
            print(msg)
        
        if not success:
            print("\n" + "="*50)
            print("ERROR: Failed to install required dependencies")
            print("="*50)
            print("\nThe packages may have been installed, but Python")
            print("cannot find them in the current session.")
            print("\nPlease try ONE of these solutions:")
            print("\n1. RESTART the application (recommended)")
            print("   Close this window and run the .bat file again")
            print("\n2. Install manually, then restart:")
            print("   pip install customtkinter Pillow openpyxl pandas requests")
            print("\n3. If using system Python on Linux:")
            print("   pip install customtkinter Pillow openpyxl pandas requests --break-system-packages")
            input("\nPress Enter to exit...")
            sys.exit(1)
        
        print("\n✓ All dependencies verified!")
        print("="*50)
        print()
    
    except Exception as e:
        print(f"Warning: Could not check dependencies: {e}")
        print("Attempting to launch anyway...")
    
    # Step 2: Now import and launch GUI
    try:
        from gui.meeting_import_app import MeetingImportGUI
        
        app = MeetingImportGUI()
        app.mainloop()
    
    except ImportError as e:
        print("\n" + "="*50)
        print("ERROR: Cannot import required modules")
        print("="*50)
        print(f"\n{e}")
        print("\nIf packages were just installed, please RESTART the application.")
        print("Close this window and run the .bat file again.")
        print("\nOr install manually:")
        print("pip install customtkinter Pillow openpyxl pandas requests")
        print("\nIf using system Python:")
        print("pip install customtkinter Pillow openpyxl pandas requests --break-system-packages")
        input("\nPress Enter to exit...")
        sys.exit(1)
    
    except Exception as e:
        try:
            from tkinter import messagebox
            messagebox.showerror("Error", f"Application error: {e}")
        except:
            print(f"Application error: {e}")
        
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
        sys.exit(1)


if __name__ == "__main__":
    main()
