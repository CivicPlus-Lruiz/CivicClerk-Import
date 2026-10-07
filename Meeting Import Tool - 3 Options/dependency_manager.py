"""
Dependency Manager Module
Handles silent installation of required Python packages
"""

import subprocess
import sys
import importlib
import site
from typing import List, Tuple


# Required packages
REQUIRED_PACKAGES = [
    ("customtkinter", "customtkinter"),
    ("PIL", "Pillow"),
    ("openpyxl", "openpyxl"),
    ("pandas", "pandas"),
    ("requests", "requests"),
    ("fitz", "PyMuPDF"),          # PDF preview and compare functionality
]

# Optional packages (large or system-dependent installs)
OPTIONAL_PACKAGES = [
    ("office365", "Office365-REST-Python-Client"),  # SharePoint integration
    ("easyocr", "easyocr"),                         # OCR filename suggestions (large download)
]


def check_package(import_name: str) -> bool:
    """Check if a package is installed."""
    try:
        importlib.import_module(import_name)
        return True
    except ImportError:
        return False


def reload_site_packages():
    """Reload site packages to pick up newly installed packages."""
    try:
        # Reload site to pick up newly installed packages
        importlib.reload(site)
        
        # Also reload sys.path
        import os
        user_site = site.getusersitepackages()
        if user_site not in sys.path:
            sys.path.insert(0, user_site)
    except:
        pass


def install_package(package_name: str) -> bool:
    """
    Install a package using pip.
    Tries without --break-system-packages first, then with it if needed.
    """
    # Try 1: Standard pip install
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", package_name, "--quiet"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        reload_site_packages()  # Refresh path after install
        return True
    except subprocess.CalledProcessError:
        pass
    
    # Try 2: With --break-system-packages (for system Python on Linux)
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", package_name, 
             "--break-system-packages", "--quiet"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        reload_site_packages()  # Refresh path after install
        return True
    except subprocess.CalledProcessError:
        pass
    
    # Try 3: With --user flag (user installation)
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", package_name, 
             "--user", "--quiet"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        reload_site_packages()  # Refresh path after install
        return True
    except subprocess.CalledProcessError:
        return False


def check_and_install_dependencies() -> Tuple[bool, List[str]]:
    """
    Check for required dependencies and install if missing.
    
    Returns:
        Tuple of (success: bool, messages: List[str])
    """
    messages = []
    all_installed = True
    packages_installed = False
    
    # Check required packages
    for import_name, package_name in REQUIRED_PACKAGES:
        if not check_package(import_name):
            messages.append(f"Installing {package_name}...")
            if install_package(package_name):
                messages.append(f"✓ Installed {package_name}")
                packages_installed = True
            else:
                messages.append(f"✗ Failed to install {package_name}")
                all_installed = False
        else:
            messages.append(f"✓ {package_name} already installed")
    
    # If we installed anything, reload site packages one more time
    if packages_installed:
        reload_site_packages()
    
    # Check optional packages (don't fail if these don't install)
    for import_name, package_name in OPTIONAL_PACKAGES:
        if not check_package(import_name):
            messages.append(f"Installing optional {package_name}...")
            if install_package(package_name):
                messages.append(f"✓ Installed {package_name}")
            else:
                messages.append(f"⚠ Optional package {package_name} not installed (SharePoint features may be limited)")
    
    # Final verification - check if packages are now importable
    if all_installed:
        messages.append("")
        messages.append("Verifying installations...")
        for import_name, package_name in REQUIRED_PACKAGES:
            if not check_package(import_name):
                messages.append(f"⚠ Warning: {package_name} installed but not importable")
                messages.append(f"   You may need to restart the application")
                all_installed = False
    
    return all_installed, messages


def get_dependency_status() -> dict:
    """Get the installation status of all dependencies."""
    status = {
        "required": {},
        "optional": {}
    }
    
    for import_name, package_name in REQUIRED_PACKAGES:
        status["required"][package_name] = check_package(import_name)
    
    for import_name, package_name in OPTIONAL_PACKAGES:
        status["optional"][package_name] = check_package(import_name)
    
    return status


if __name__ == "__main__":
    # Test the dependency checker
    print("Checking dependencies...")
    success, messages = check_and_install_dependencies()
    
    for msg in messages:
        print(msg)
    
    if success:
        print("\n✓ All required dependencies are installed!")
    else:
        print("\n✗ Some dependencies failed to install or import")
        print("Please install manually using:")
        print("pip install customtkinter Pillow openpyxl pandas requests PyMuPDF")
        print("\nThen restart the application.")
