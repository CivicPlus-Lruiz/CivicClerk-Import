"""
Backend Wrapper Module
Interfaces between GUI and the existing Meeting_Import script
"""

import os
import sys
import json
import shutil
import subprocess
from pathlib import Path
from typing import Optional, Dict, Any, List
import pandas as pd


class BackendWrapper:
    """Wraps the existing backend script for GUI integration."""
    
    def __init__(self, script_path: Optional[str] = None):
        self.script_path = script_path
        
        if not script_path or not Path(script_path).exists():
            print("⚠ Backend script not found")
    
    def run_import(self, config: dict) -> Dict[str, Any]:
        """
        Execute the import process with GUI configuration.
        
        Args:
            config: {
                "base_dir": str,
                "date_format": str,
                "dry_run": bool,
                "meeting_bodies": list[dict],
                "subtypes": list[dict]
            }
        
        Returns:
            dict: {
                "records": list[dict],
                "warnings": list[str],
                "duplicates": list[str],
                "errors": list[dict]
            }
        """
        try:
            # Prepare the import environment
            base_dir = config["base_dir"]
            dry_run = config.get("dry_run", False)
            print(f"\n{'='*60}")
            print(f"BACKEND WRAPPER: Running import (dry_run={dry_run})")
            print(f"Base directory: {base_dir}")
            print(f"{'='*60}\n")
            
            self._prepare_excel_config(config, base_dir)
            
            # Run the backend script
            result = self._execute_backend_script(base_dir, config)
            
            # Parse outputs
            output_data = self._parse_output_files(base_dir, config)
            
            return output_data
        
        except Exception as e:
            return {
                "records": [],
                "warnings": [],
                "duplicates": [],
                "errors": [
                    {
                        "error_type": "System Error",
                        "file_name": "N/A",
                        "issue": str(e),
                        "original_location": "",
                        "current_location": "",
                        "file_path": ""
                    }
                ]
            }
    
    def _prepare_excel_config(self, config: dict, base_dir: str):
        """Create Import Sheet.xlsx from GUI configuration."""
        try:
            import openpyxl
            from openpyxl import Workbook
            
            wb = Workbook()
            
            # Create Meeting Bodies sheet
            ws_bodies = wb.active
            ws_bodies.title = "Meeting Bodies"
            
            # Headers - MUST match what backend script expects
            # Adding BOTH column names to handle different backend versions
            ws_bodies['A1'] = "Folder Path"
            ws_bodies['B1'] = "Meeting Type"
            ws_bodies['C1'] = "Event Category"
            ws_bodies['D1'] = "Default Time"
            ws_bodies['E1'] = "Default Time Text"
            ws_bodies['F1'] = "Date Format"
            
            # Data
            for idx, body in enumerate(config["meeting_bodies"], start=2):
                ws_bodies[f'A{idx}'] = body.get("folder_path", "")
                ws_bodies[f'B{idx}'] = body.get("meeting_type", "")
                ws_bodies[f'C{idx}'] = body.get("event_category", "")
                ws_bodies[f'D{idx}'] = body.get("default_time", "")
                ws_bodies[f'E{idx}'] = body.get("default_time", "")  # Same value in both columns
                ws_bodies[f'F{idx}'] = config.get("date_format", "US DATE")
            
            # Create Subtypes sheet
            ws_subtypes = wb.create_sheet("Subtypes")
            
            # Headers - Adding both column name variants
            ws_subtypes['A1'] = "Event Category"
            ws_subtypes['B1'] = "Folder Name (optional)"
            ws_subtypes['C1'] = "Subtype"
            ws_subtypes['D1'] = "Aliases"
            ws_subtypes['E1'] = "Override Time"
            ws_subtypes['F1'] = "Default Time Text"  # Backend may expect this name
            ws_subtypes['G1'] = "Override Meeting Type"
            ws_subtypes['H1'] = "Event Category Override"
            
            # Data
            for idx, subtype in enumerate(config["subtypes"], start=2):
                ws_subtypes[f'A{idx}'] = subtype.get("event_category", "")
                ws_subtypes[f'B{idx}'] = subtype.get("folder_name", "")
                ws_subtypes[f'C{idx}'] = subtype.get("subtype", "")
                ws_subtypes[f'D{idx}'] = subtype.get("aliases", "")
                ws_subtypes[f'E{idx}'] = subtype.get("override_time", "")
                ws_subtypes[f'F{idx}'] = subtype.get("override_time", "")  # Same value for both time columns
                ws_subtypes[f'G{idx}'] = subtype.get("override_meeting_type", "")
                ws_subtypes[f'H{idx}'] = subtype.get("event_category_override", "")
            
            # Save to base directory
            excel_path = Path(base_dir) / "Import Sheet.xlsx"
            wb.save(excel_path)
            
            print(f"✓ Created Import Sheet.xlsx")
            
            # Verify what was created
            try:
                df_check = pd.read_excel(excel_path, sheet_name="Meeting Bodies")
                print(f"  Column headers created: {list(df_check.columns)}")
            except Exception as e:
                print(f"  Warning: Could not verify columns: {e}")
        
        except Exception as e:
            raise Exception(f"Failed to create Excel config: {str(e)}")
    
    def _execute_backend_script(self, base_dir: str, config: dict = None) -> bool:
        """Execute the backend Python script."""
        if not self.script_path or not Path(self.script_path).exists():
            raise Exception("Backend script not available")
        
        try:
            # Copy script to base directory
            script_dest = Path(base_dir) / Path(self.script_path).name
            shutil.copy(self.script_path, script_dest)
            
            # Execute script with optional dry-run flag
            cmd = [sys.executable, str(script_dest)]
            if config.get("dry_run", False):
                cmd.append("--dry-run")
            
            result = subprocess.run(
                cmd,
                cwd=base_dir,
                capture_output=True,
                text=True,
                timeout=300  # 5 minute timeout
            )
            
            if result.returncode != 0:
                print(f"⚠ Script execution had errors:\n{result.stderr}")
            
            print(f"✓ Script executed")
            print(result.stdout)
            
            return True
        
        except subprocess.TimeoutExpired:
            raise Exception("Script execution timed out (>5 minutes)")
        except Exception as e:
            raise Exception(f"Failed to execute script: {str(e)}")
    
    def _parse_output_files(self, base_dir: str, config: dict = None) -> Dict[str, Any]:
        """Parse output files created by the backend script."""
        result = {
            "records": [],
            "warnings": [],
            "duplicates": [],
            "errors": []
        }
        
        if config is None:
            config = {}
        
        print(f"Parsing output files from: {base_dir}")
        
        try:
            # Parse Meetings_Output.xlsx (check both regular and DryRun_ versions)
            dry_run = config.get("dry_run", False)
            output_filename = "DryRun_Meetings_Output.xlsx" if dry_run else "Meetings_Output.xlsx"
            output_file = Path(base_dir) / output_filename
            
            # Fallback to regular name if DryRun_ not found
            if not output_file.exists() and dry_run:
                output_file = Path(base_dir) / "Meetings_Output.xlsx"
            
            if output_file.exists():
                print(f"Reading output file: {output_file.name}")
                df = pd.read_excel(output_file)
                result["records"] = df.to_dict('records')
                print(f"✓ Loaded {len(result['records'])} records")
            else:
                print(f"⚠ Output file not found: {output_filename}")
            
            # Parse Error_Log.csv (check both regular and DryRun_ versions)
            error_filename = "DryRun_Error_Log.csv" if dry_run else "Error_Log.csv"
            error_file = Path(base_dir) / error_filename
            
            # Fallback to regular name if DryRun_ not found
            if not error_file.exists() and dry_run:
                error_file = Path(base_dir) / "Error_Log.csv"
            
            print(f"Checking for error file: {error_file}")
            if error_file.exists():
                print(f"✓ {error_file.name} found, reading...")
                df_errors = pd.read_csv(error_file)
                print(f"  Found {len(df_errors)} error rows")
                
                for _, row in df_errors.iterrows():
                    # Determine error type from destination path
                    dest = str(row.get("Destination", ""))
                    error_type = "Unknown"
                    
                    if "Ambiguous_Dates" in dest:
                        error_type = "Ambiguous Date"
                    elif "Weekend_Dates" in dest:
                        error_type = "Weekend Date"
                    elif "Duplicates" in dest:
                        error_type = "Duplicate"
                    elif "Unrecognized_Format" in dest:
                        error_type = "Unrecognized Format"
                    
                    # Extract original location from Review structure
                    # Review path: Review_Files/<ErrorType>/<Body>/<SourceFolder>/filename
                    # Original path: meetings/<Body>/<SourceFolder>/filename
                    original_loc = ""
                    if dest:
                        dest_path = Path(dest)
                        try:
                            parts = list(dest_path.parts)
                            if "Review_Files" in parts:
                                idx = parts.index("Review_Files")
                                # FIX: use local variable 'dest_base' instead of
                                # shadowing the outer 'base_dir' parameter
                                dest_base = Path(*parts[:idx])
                                
                                # Get the parts after Review_Files/<ErrorType>/
                                # which are: <Body>/<SourceFolder>/filename
                                after_review = parts[idx+2:]  # Skip Review_Files and ErrorType
                                
                                if after_review:
                                    # Remove the filename to get just the directory
                                    body_and_folder = after_review[:-1]  # All but filename
                                    
                                    # Reconstruct: base/meetings/<Body>/<SourceFolder>
                                    original_loc = str(dest_base / "meetings" / Path(*body_and_folder))
                        except:
                            # Fallback to just /meetings if parsing fails
                            if "Review_Files" in parts:
                                idx = parts.index("Review_Files")
                                original_loc = str(Path(*parts[:idx]) / "meetings")
                    
                    # During dry run, files were never moved - use original location
                    file_name = str(row.get("File", ""))
                    if dry_run:
                        actual_location = original_loc
                        actual_file_path = str(Path(original_loc) / file_name) if original_loc else dest
                    else:
                        actual_location = dest
                        actual_file_path = dest

                    error_dict = {
                        "error_type": error_type,
                        "file_name": file_name,
                        "issue": str(row.get("Issue", "")),
                        "original_location": original_loc,
                        "current_location": actual_location,
                        "file_path": actual_file_path
                    }
                    
                    result["errors"].append(error_dict)
                
                print(f"✓ Parsed {len(result['errors'])} errors from Error_Log.csv")
            else:
                print(f"⚠ Error_Log.csv not found at: {error_file}")
            
            # Scan Review folders for additional info
            # FIX: base_dir is now correctly preserved (no longer shadowed above)
            review_dir = Path(base_dir) / "Review_Files"
            if review_dir.exists():
                # Check Duplicates folder
                dup_dir = review_dir / "Duplicates"
                if dup_dir.exists():
                    for file in dup_dir.rglob("*"):
                        if file.is_file():
                            result["duplicates"].append(f"Duplicate: {file.name}")
                
                # Could add warnings for other Review folders
                for subfolder in ["Ambiguous_Dates", "Weekend_Dates", "Unrecognized_Format"]:
                    folder_path = review_dir / subfolder
                    if folder_path.exists():
                        file_count = len(list(folder_path.rglob("*.*")))
                        if file_count > 0:
                            result["warnings"].append(
                                f"{file_count} files in {subfolder} require review"
                            )
        
        except Exception as e:
            result["warnings"].append(f"Error parsing output files: {str(e)}")
        
        return result
    
    def get_script_info(self) -> Dict[str, str]:
        """Get information about the backend script."""
        if self.script_path and Path(self.script_path).exists():
            return {
                "path": self.script_path,
                "name": Path(self.script_path).name,
                "size": f"{Path(self.script_path).stat().st_size / 1024:.1f} KB",
                "exists": "Yes"
            }
        else:
            return {
                "path": "N/A",
                "name": "N/A",
                "size": "N/A",
                "exists": "No"
            }
