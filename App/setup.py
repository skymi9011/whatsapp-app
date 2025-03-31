import os
import subprocess
import sys
import tempfile
import shutil
import urllib.request

# Maximum number of retries for commands
MAX_RETRIES = 2

# Use the user's temporary directory for the log file
LOG_FILE = ".\\setup.log"

def log_message(message):
    """Log messages to a file and print them to the console."""
    with open(LOG_FILE, "a") as log:
        log.write(message + "\n")
    print(message)

def create_shortcut():
    """Create a shortcut for the application on the desktop."""
    desktop = winshell.desktop()  # مسار سطح المكتب
    shortcut_path = os.path.join(desktop, "WhatsAppSender.lnk")  # اسم الاختصار
    target = os.path.join(os.getcwd(), "WhatsAppSender.exe")  # مسار التطبيق
    with winshell.shortcut(shortcut_path) as shortcut:
        shortcut.path = target
        shortcut.description = "WhatsApp Sender Application"

    log_message(f">> Shortcut created at: {shortcut_path}")

def download_get_pip():
    """Download get-pip.py using urllib if not already downloaded."""
    url = "https://bootstrap.pypa.io/get-pip.py"
    file_name = "get-pip.py"
    if not os.path.exists(file_name):
        log_message(f">> Downloading {file_name} from {url}...")
        try:
            with urllib.request.urlopen(url) as response:
                with open(file_name, "wb") as file:
                    file.write(response.read())
            log_message(f">> Successfully downloaded {file_name}.")
        except Exception as e:
            log_message(f">> Failed to download {file_name}: {e}")
            sys.exit(1)


def run_command(command, retries=MAX_RETRIES, next_step=False):
    
    """Run system commands with retries and handle missing modules."""
    try:
        result = subprocess.run(command, shell=True, text=True, capture_output=True)
        if command == 'python -m venv venv' and os.path.exists(".\\venv"):
            return 0  
        if result.returncode != 0:
            log_message(result.stderr.strip())
            # Check for "No module named" error
            if "No module named" in result.stderr:
                missing_module = result.stderr.split("No module named")[-1].strip().strip("'")
                log_message(f">> Missing module detected: {missing_module}. Installing...")
                run_command(f'.\\venv\\Scripts\\python.exe -m pip install {missing_module}')
                # Retry the original command after installing the module
                return run_command(command, retries - 1, next_step)
            elif retries > 0:
                log_message(f"Error while executing the command: {command}. Retries left: {retries - 1}")
                return run_command(command, retries - 1, next_step)
            else:
                log_message(f"Command failed after retries: {command}")
                if next_step:
                    return MAX_RETRIES  # Reset retries for the next step
                else:
                    sys.exit(1)
    except Exception as e:
        log_message(f"Command failed with error: {e}")
        sys.exit(1)

def main():
    log_message("=== Starting the application setup ===")

    # 1. Create a virtual environment
    log_message(">> Creating a virtual environment...")
    if not os.path.exists(".\\venv"):
        run_command(f'python -m venv venv')
    else:
        log_message(">> Virtual environment already exists.")

    # 2. Install pip if not found
    if not os.path.exists(".\\venv\\Scripts\\pip.exe"):
        log_message(">> Pip not found, installing manually...")
        if not os.path.exists("get-pip.py"):
            log_message(">> Downloading get-pip.py...")
            download_get_pip()
            run_command(f'.\\venv\\Scripts\\python.exe get-pip.py')

        if not os.path.exists("get-pip.py"):
            log_message(">> Failed to download get-pip.py. Please download it manually from https://bootstrap.pypa.io/get-pip.py")
            sys.exit(1)

        run_command(f'.\\venv\\Scripts\\python.exe get-pip.py')

    # Verify pip installation
    if not os.path.exists(".\\venv\\Scripts\\pip.exe"):
        log_message(">> Pip still not found, attempting to upgrade manually...")
        run_command(f'.\\venv\\Scripts\\python.exe -m pip install --upgrade pip')

    # 3. Activate the virtual environment
    log_message(">> Activating the virtual environment...")
    os.environ["VIRTUAL_ENV"] = os.path.abspath("venv")
    os.environ["PATH"] = os.path.join(os.environ["VIRTUAL_ENV"], "Scripts") + os.pathsep + os.environ["PATH"]
    python = f'"{os.path.join(os.environ["VIRTUAL_ENV"], "Scripts", "python.exe")}" -m '
    pyinstaller = f'"{os.path.join(os.environ["VIRTUAL_ENV"], "Scripts", "pyinstaller.exe")}" '

    # 4. Install dependencies
    log_message(">> Installing dependencies...")
    run_command(python + "pip install --upgrade pip")
    run_command(python + "pip install pyinstaller pandas selenium PyQt5 requests openpyxl")

    # 5. Create an executable using PyInstaller
    
    log_message(">> Creating an executable...")
    run_command(pyinstaller + '--onedir --windowed --icon=app.ico --name "WhatsAppSender" app.py')
    dist_dir = ".\\dist\\WhatsAppSender"
    main_dir = os.getcwd()
    
    source = os.path.join(main_dir, "app.ico")
    destination = os.path.join(dist_dir, "app.ico")
    shutil.move(source, destination)
    if os.path.exists(dist_dir):
        for item in os.listdir(dist_dir):
            source = os.path.join(dist_dir, item)
            destination = os.path.join(main_dir, item)
            if os.path.isdir(source):
                shutil.move(source, destination)
            else:
                shutil.copy2(source, destination)
    
    # Clean up unnecessary directories and files
    shutil.rmtree('.\\venv', ignore_errors=True)
    shutil.rmtree('.\\build', ignore_errors=True)
    shutil.rmtree('.\\dist', ignore_errors=True)
    for file in os.listdir(main_dir):
        if file.endswith(".spec") or file.endswith(".py"):
            os.remove(os.path.join(main_dir, file))
        
    log_message("=== Installation Complete ! ===")
   
    log_message(f">> Log file saved at: {LOG_FILE}")

def new_func():
    try:
        __import__("winshell")
    except ImportError:
        print("Installing winshell...")
        subprocess.check_call([".\\venv\\Scripts\\python.exe", "-m", "pip", "install", "winshell"])
        print("winshell installed successfully!")
    finally:
        globals()["winshell"] = __import__("winshell")
    try:
        __import__("win32com")
    except ImportError:
        print(f"Installing win32com...")
        subprocess.check_call([".\\venv\\Scripts\\python.exe", "-m", "pip", "install", "pywin32"])
        print(f"win32com installed successfully!")
    finally:
        globals()["win32com"] = __import__("win32com")

if __name__ == "__main__":
    main()
    new_func()
    import winshell # type: ignore
    create_shortcut()