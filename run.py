import sys
import os
import subprocess
import time
import webbrowser
import threading

def check_dependencies():
    print("Checking dependencies...")
    required_modules = ['flask', 'flask_cors', 'docx', 'docxtpl']
    missing = False
    for module in required_modules:
        try:
            # docx is imported as docx, docxtpl as docxtpl
            if module == 'docx':
                import docx
            elif module == 'docxtpl':
                import docxtpl
            else:
                __import__(module)
        except ImportError:
            missing = True
            print(f"  - Missing package: {module}")
            break
    
    if missing:
        print("Installing missing dependencies from requirements.txt...")
        try:
            # Run pip module to install
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "backend/requirements.txt"])
            print("Dependencies installed successfully!\n")
        except Exception as e:
            print(f"Error installing dependencies automatically: {e}")
            print("Please run the command manually: python -m pip install -r backend/requirements.txt")
            sys.exit(1)
    else:
        print("All dependencies are satisfied!\n")

def open_browser():
    # Wait for the Flask server to initialize
    time.sleep(1.5)
    print("=" * 60)
    print("[INFO] WordDoc is ready!")
    print("URL: http://localhost:5000")
    print("Paste your screenshots and text in the browser to start creating!")
    print("=" * 60)

    webbrowser.open('http://localhost:5000')

if __name__ == '__main__':
    print("=" * 60)
    print("          WordDoc - Instant Word Creator Launcher")
    print("=" * 60)
    
    # Force working directory to the directory of run.py
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    
    # Check dependencies
    check_dependencies()
    
    # Start browser thread
    browser_thread = threading.Thread(target=open_browser)
    browser_thread.daemon = True
    browser_thread.start()
    
    # Start backend server
    print("Starting local server...")
    sys.path.append(os.path.join(script_dir, 'backend'))
    try:
        from app import app
        # Run Flask server (debug=False to disable file watchers and clean run in threads)
        app.run(port=5000, debug=False, host='127.0.0.1')
    except Exception as e:
        print(f"\nServer failed to start: {e}")
        print("Please check if port 5000 is already in use by another application.")
        input("\nPress Enter to exit...")
