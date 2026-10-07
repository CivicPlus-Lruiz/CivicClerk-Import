"""
SharePoint Integration Module
Handles downloading and caching the latest backend script from SharePoint
"""

import os
import re
import shutil
from pathlib import Path
from typing import Optional
from datetime import datetime, timedelta
import requests

# Optional Office365 integration (for direct SharePoint API access)
try:
    from office365.sharepoint.client_context import ClientContext
    from office365.runtime.auth.authentication_context import AuthenticationContext
    HAS_OFFICE365 = True
except ImportError:
    HAS_OFFICE365 = False
    ClientContext = None
    AuthenticationContext = None


class SharePointManager:
    """Manages SharePoint integration for script updates."""
    
    # Multiple possible OneDrive sync paths (check in order)
    SCRIPT_PATHS = [
        # Standard OneDrive path
        Path.home() / "OneDrive - CivicPlus" / "Documents - CP Oregon - Web + Meetings Team" / "Operations - Meetings" / "_Luis" / "Imports Folder" / "Latest Script",
        # Alternative path (some users have different sync structure)
        Path.home() / "CivicPlus" / "CP Oregon - Web + Meetings Team - Documents" / "Operations - Meetings" / "_Luis" / "Imports Folder" / "Latest Script"
    ]
    
    SCRIPT_PATTERN = r"Meeting_Import_Stable_v(\d+)\.(\d+)\.(\d+)\.py"  # v3.5.3 format with dots
    
    def __init__(self):
        self.cache_dir = Path.home() / ".meeting_import_cache"
        self.cache_dir.mkdir(exist_ok=True)
        
        self.version_file = self.cache_dir / "version.txt"
        self.cached_script = None
        self.current_version = None
        self.script_folder = None  # Will be set when we find the folder
        
        self._load_cached_version()
    
    def _load_cached_version(self):
        """Load information about cached script."""
        if self.version_file.exists():
            try:
                with open(self.version_file, 'r') as f:
                    lines = f.readlines()
                    if len(lines) >= 2:
                        self.current_version = lines[0].strip()
                        script_name = lines[1].strip()
                        script_path = self.cache_dir / script_name
                        if script_path.exists():
                            self.cached_script = str(script_path)
            except Exception:
                pass
    
    def get_latest_script(self) -> Optional[str]:
        """
        Get the latest script from OneDrive synced SharePoint folder.
        Checks multiple possible paths to support different sync configurations.
        
        Returns:
            str: Path to the script file, or None if unavailable
        """
        # Try each possible OneDrive path
        for idx, script_folder in enumerate(self.SCRIPT_PATHS):
            print(f"Checking OneDrive path {idx + 1}/{len(self.SCRIPT_PATHS)}: {script_folder}")
            
            if script_folder.exists():
                print(f"✓ Folder exists")
                self.script_folder = script_folder  # Save the working path
                
                # List all files in the folder for debugging
                try:
                    all_files = list(script_folder.iterdir())
                    print(f"  Found {len(all_files)} files/folders:")
                    for file in all_files[:10]:  # Show first 10
                        print(f"    - {file.name}")
                except Exception as e:
                    print(f"  Could not list folder contents: {e}")
                
                # Find all matching scripts
                script_files = []
                for file in script_folder.glob("*.py"):
                    print(f"  Checking Python file: {file.name}")
                    match = re.match(self.SCRIPT_PATTERN, file.name)
                    if match:
                        version = tuple(map(int, match.groups()))
                        script_files.append((version, str(file), file.name))
                        print(f"    ✓ Matches pattern! Version: {version}")
                    else:
                        print(f"    ✗ Does not match pattern: {self.SCRIPT_PATTERN}")
                
                if script_files:
                    # Sort by version and get latest
                    script_files.sort(reverse=True, key=lambda x: x[0])
                    latest_version, latest_path, latest_name = script_files[0]
                    version_str = f"v{latest_version[0]}.{latest_version[1]}.{latest_version[2]}"
                    
                    print(f"✓ Found script in OneDrive: {latest_name} ({version_str})")
                    
                    # Cache it for offline use
                    self._cache_script(latest_path, version_str, latest_name)
                    
                    return latest_path
                else:
                    print(f"⚠ No scripts found in this folder matching pattern: {self.SCRIPT_PATTERN}")
            else:
                print(f"✗ Folder does not exist")
        
        # No working path found
        print(f"⚠ No OneDrive folder found in any of the checked locations")
        print("  Make sure OneDrive is syncing the SharePoint folder")
        
        # FALLBACK: Check local GUI directory for script
        print("\nChecking local GUI directory for fallback script...")
        try:
            # Get the directory where sharepoint_integration.py is located
            local_dir = Path(__file__).parent
            
            # Look for script files in the same directory
            script_files = []
            for file in local_dir.glob("*.py"):
                match = re.match(self.SCRIPT_PATTERN, file.name)
                if match:
                    version = tuple(map(int, match.groups()))
                    script_files.append((version, str(file), file.name))
                    print(f"  Found local script: {file.name} (v{version[0]}.{version[1]}.{version[2]})")
            
            if script_files:
                # Sort by version and get latest
                script_files.sort(reverse=True, key=lambda x: x[0])
                latest_version, latest_path, latest_name = script_files[0]
                version_str = f"v{latest_version[0]}.{latest_version[1]}.{latest_version[2]}"
                
                print(f"✓ Using local script: {latest_name} ({version_str})")
                
                # Cache it
                self._cache_script(latest_path, version_str, latest_name)
                
                return latest_path
            else:
                print("  No local script found matching pattern")
        
        except Exception as e:
            print(f"  Error checking local directory: {e}")
        
        # Fall back to cached version
        if self.cached_script and Path(self.cached_script).exists():
            print(f"✓ Using cached script: {Path(self.cached_script).name} ({self.current_version})")
            return self.cached_script
        
        print("⚠ No script available")
        print("  Checked paths:")
        for path in self.SCRIPT_PATHS:
            print(f"    - {path}")
        print("  Also checked: Local GUI directory")
        print("\n  TIP: You can drop the backend script directly in the GUI folder")
        print("       (same folder as sharepoint_integration.py)")
        print("       Script must match pattern: Meeting_Import_Stable_vX.X.X.py")
        return None
    
    def _cache_script(self, script_path: str, version_str: str, script_name: str):
        """Cache a script file for offline use."""
        try:
            import shutil
            
            # Copy to cache
            dest_path = self.cache_dir / script_name
            shutil.copy2(script_path, dest_path)
            
            # Update version file
            with open(self.version_file, 'w') as f:
                f.write(f"{version_str}\n")
                f.write(f"{script_name}\n")
                f.write(f"{datetime.now().isoformat()}\n")
            
            # Update cached info
            self.current_version = version_str
            self.cached_script = str(dest_path)
        
        except Exception:
            pass  # Silent fail on caching
    
    def get_version_info(self) -> dict:
        """Get information about current cached version."""
        return {
            "version": self.current_version,
            "script_path": self.cached_script,
            "cache_dir": str(self.cache_dir)
        }
    
    def clear_cache(self):
        """Clear cached scripts."""
        try:
            if self.cache_dir.exists():
                shutil.rmtree(self.cache_dir)
            self.cache_dir.mkdir(exist_ok=True)
            self.current_version = None
            self.cached_script = None
            return True
        except Exception as e:
            print(f"⚠ Failed to clear cache: {str(e)}")
            return False


# Alternative implementation using direct HTTP if Office365 library not available
class SharePointManagerSimple:
    """Simplified SharePoint manager using direct download links."""
    
    def __init__(self):
        self.cache_dir = Path.home() / ".meeting_import_cache"
        self.cache_dir.mkdir(exist_ok=True)
        
        self.version_file = self.cache_dir / "version.txt"
        self.cached_script = None
        self.current_version = None
        
        # This would need to be a direct download link
        # Example: https://civicplus.sharepoint.com/:u:/s/Site/EaBcDeFgHiJkLmNoPqRsTuVw
        self.download_url = None
        
        self._load_cached_version()
    
    def _load_cached_version(self):
        """Load information about cached script."""
        if self.version_file.exists():
            try:
                with open(self.version_file, 'r') as f:
                    lines = f.readlines()
                    if len(lines) >= 2:
                        self.current_version = lines[0].strip()
                        script_name = lines[1].strip()
                        script_path = self.cache_dir / script_name
                        if script_path.exists():
                            self.cached_script = str(script_path)
            except Exception:
                pass
    
    def get_latest_script(self) -> Optional[str]:
        """Get the latest script."""
        # For now, just return cached version
        # Full implementation would need a direct download link
        if self.cached_script and Path(self.cached_script).exists():
            return self.cached_script
        
        # Could implement manual download instructions here
        return None
    
    def get_version_info(self) -> dict:
        """Get version info."""
        return {
            "version": self.current_version,
            "script_path": self.cached_script,
            "cache_dir": str(self.cache_dir)
        }


# Use the appropriate manager based on available libraries
try:
    from office365.sharepoint.client_context import ClientContext
    SharePointManagerClass = SharePointManager
except ImportError:
    print("⚠ Office365 library not available, using simplified manager")
    SharePointManagerClass = SharePointManagerSimple
