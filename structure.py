import os
import argparse
import shutil
from pathlib import Path
from shutil import move, rmtree, copytree
from datetime import datetime
import sys

class DataOrganizer:
    def __init__(self, dry_run=False, create_backup=True, verbose=False):
        self.dry_run = dry_run
        self.create_backup = create_backup
        self.verbose = verbose
        self.stats = {
            'files_moved': 0,
            'folders_deleted': 0,
            'already_organized': 0,
            'directories_processed': 0,
            'skipped_exists': 0,
            'skipped_errors': 0
        }
        self.operations = {'moves': [], 'deletions': [], 'organized': [], 'warnings': []}
        
    def check_environment(self):
        """Validate we're in the right directory and have proper permissions"""
        print("🔍 Checking environment...")
        
        # Check if we're in the right place
        root = Path(".")
        mice_folders = list(root.glob("Mice_*")) + list(root.glob("PV_Mice*"))
        
        if not mice_folders:
            print("❌ No Mice_* or PV_Mice* folders found in current directory.")
            print("💡 Make sure you're in the directory containing your mouse data folders.")
            print("   Expected structure: ./Mice_001/, ./Mice_002/, etc.")
            return False
            
        print(f"✅ Found {len(mice_folders)} mouse data folders - looks good!")
        
        # Check write permissions
        if not os.access(".", os.W_OK):
            print("❌ No write permissions in current directory!")
            return False
            
        # Check disk space (basic check)
        try:
            total, used, free = shutil.disk_usage(".")
            free_gb = free // (1024**3)
            if free_gb < 1:
                print(f"⚠️  Low disk space: only {free_gb}GB available")
            else:
                print(f"✅ Disk space OK: {free_gb}GB available")
        except:
            print("⚠️  Could not check disk space")
            
        return True
    
    def explain_purpose(self):
        """Explain what this script does"""
        print("\n🎯 PNN Data Organization Script")
        print("This prepares your microscopy images for the counting pipeline.")
        print("Each .tif image needs to be in its own folder for batch processing.")
        print("\nExample transformation:")
        print("  BEFORE: Mice_001/Mouse_01/WFA_1L.tif")
        print("  AFTER:  Mice_001/Mouse_01/WFA_1L/WFA_1L.tif")
        
    def scan_data_structure(self):
        """Analyze current data organization"""
        print("\n🔍 Scanning your data structure...")
        root = Path(".")
        
        for mice_dir in root.iterdir():
            if not mice_dir.is_dir():
                continue
            if not (mice_dir.name.startswith("Mice_") or mice_dir.name.startswith("PV_Mice")):
                continue
                
            if self.verbose:
                print(f"  📁 Scanning {mice_dir.name}...")
                
            for mouse_dir in mice_dir.glob("Mouse_*"):
                self.stats['directories_processed'] += 1
                
                try:
                    items = list(mouse_dir.iterdir())
                except PermissionError as e:
                    print(f"  ⚠️  Permission denied: {mouse_dir}")
                    self.operations['warnings'].append(f"Permission denied: {mouse_dir}")
                    self.stats['skipped_errors'] += 1
                    continue
                except Exception as e:
                    print(f"  ⚠️  Error scanning {mouse_dir}: {e}")
                    self.operations['warnings'].append(f"Error scanning {mouse_dir}: {e}")
                    self.stats['skipped_errors'] += 1
                    continue
                
                for item in items:
                    try:
                        # Skip symbolic links
                        if item.is_symlink():
                            if self.verbose:
                                print(f"  ⚠️  Skipping symbolic link: {item}")
                            continue
                            
                        if item.is_file() and item.suffix.lower() == ".tif":
                            # Loose .tif file that needs organizing
                            folder_path = mouse_dir / item.stem
                            dest_file = folder_path / item.name
                            
                            # Check if destination already exists
                            if dest_file.exists():
                                self.operations['warnings'].append(f"File already exists: {dest_file}")
                                self.stats['skipped_exists'] += 1
                                if self.verbose:
                                    print(f"  ⚠️  Skipping {item.name} - already exists at destination")
                                continue
                                
                            self.operations['moves'].append({
                                'source': item,
                                'dest_folder': folder_path,
                                'dest_file': dest_file
                            })
                            
                        elif item.is_dir():
                            # Check if this is a properly organized folder
                            try:
                                tif_inside = any(f.suffix.lower() == ".tif" for f in item.iterdir())
                            except PermissionError:
                                if self.verbose:
                                    print(f"  ⚠️  Cannot read folder: {item}")
                                continue
                                
                            if tif_inside:
                                self.operations['organized'].append(item)
                                continue
                                
                            # Check if matching .tif exists outside (with case variations)
                            tif_outside = None
                            for ext in ['.tif', '.TIF', '.Tif']:
                                potential_file = mouse_dir / f"{item.name}{ext}"
                                if potential_file.exists():
                                    tif_outside = potential_file
                                    break
                                    
                            if tif_outside:
                                dest_file = item / tif_outside.name
                                if dest_file.exists():
                                    self.operations['warnings'].append(f"File already exists: {dest_file}")
                                    self.stats['skipped_exists'] += 1
                                    continue
                                    
                                self.operations['moves'].append({
                                    'source': tif_outside,
                                    'dest_folder': item,
                                    'dest_file': dest_file
                                })
                                continue
                                
                            # Check if folder should be deleted - enhanced safety
                            try:
                                folder_contents = list(item.iterdir())
                                
                                # Skip if folder has any CSV or prediction files
                                has_csv = any(f.name.endswith(".csv") for f in folder_contents if f.is_file())
                                has_preds = any("_predictions" in f.name for f in folder_contents if f.is_file())
                                
                                if has_csv or has_preds:
                                    continue
                                    
                                # Only delete if empty or contains only safe file types
                                if folder_contents:  # Not empty
                                    safe_extensions = {'.txt', '.log', '.tmp', '.temp'}
                                    has_unsafe = any(
                                        f.is_file() and f.suffix.lower() not in safe_extensions 
                                        for f in folder_contents
                                    )
                                    if has_unsafe:
                                        if self.verbose:
                                            print(f"  ℹ️  Keeping {item.name} - contains other files")
                                        continue
                                        
                                self.operations['deletions'].append(item)
                                
                            except PermissionError:
                                if self.verbose:
                                    print(f"  ⚠️  Cannot check folder contents: {item}")
                                continue
                                
                    except Exception as e:
                        print(f"  ⚠️  Error processing {item}: {e}")
                        self.stats['skipped_errors'] += 1
                        continue
                        
        # Update stats
        self.stats['files_moved'] = len(self.operations['moves'])
        self.stats['folders_deleted'] = len(self.operations['deletions'])
        self.stats['already_organized'] = len(self.operations['organized'])
        
    def show_analysis_summary(self):
        """Show what was found during scanning"""
        print("\n📊 Analysis complete:")
        print(f"   - {self.stats['already_organized']} images already organized ✅")
        print(f"   - {self.stats['files_moved']} loose .tif files need organizing 📦")
        print(f"   - {self.stats['folders_deleted']} empty folders will be deleted 🗑️")
        print(f"   - {self.stats['directories_processed']} mouse directories scanned")
        
        if self.stats['skipped_exists'] > 0:
            print(f"   - {self.stats['skipped_exists']} files skipped (already exist) ⚠️")
        if self.stats['skipped_errors'] > 0:
            print(f"   - {self.stats['skipped_errors']} items skipped due to errors ⚠️")
            
        if self.operations['warnings']:
            print(f"\n⚠️  {len(self.operations['warnings'])} warnings detected. Show details? (y/n): ", end="")
            if input().strip().lower() == 'y':
                print("\nWarnings:")
                for warning in self.operations['warnings'][:10]:
                    print(f"   - {warning}")
                if len(self.operations['warnings']) > 10:
                    print(f"   ... and {len(self.operations['warnings']) - 10} more warnings")
                    
        if self.stats['files_moved'] == 0 and self.stats['folders_deleted'] == 0:
            print("\n✅ Everything already looks organized! Nothing to do.")
            return False
            
        return True
        
    def show_detailed_preview(self):
        """Show exactly what operations will be performed"""
        print("\n📋 PREVIEW - Here's exactly what I'll do:")
        
        if self.operations['moves']:
            print(f"\nMOVES ({len(self.operations['moves'])}):")
            for i, op in enumerate(self.operations['moves'][:10]):  # Show first 10
                rel_source = op['source'].relative_to(Path("."))
                rel_dest = op['dest_file'].relative_to(Path("."))
                print(f"   📦 {rel_source} → {rel_dest}")
            if len(self.operations['moves']) > 10:
                print(f"   ... and {len(self.operations['moves']) - 10} more files")
                
        if self.operations['deletions']:
            print(f"\nDELETIONS ({len(self.operations['deletions'])}):")
            for folder in self.operations['deletions'][:5]:  # Show first 5
                rel_path = folder.relative_to(Path("."))
                print(f"   🗑️ {rel_path} (no images, CSVs, or predictions)")
            if len(self.operations['deletions']) > 5:
                print(f"   ... and {len(self.operations['deletions']) - 5} more folders")
                
        if self.operations['organized']:
            print(f"\nSKIP - Already Organized ({len(self.operations['organized'])}):")
            for i, folder in enumerate(self.operations['organized'][:5]):  # Show first 5
                rel_path = folder.relative_to(Path("."))
                try:
                    tif_files = [f for f in folder.iterdir() if f.suffix.lower() == ".tif"]
                    if tif_files:
                        print(f"   ✅ {rel_path}/ (contains {tif_files[0].name})")
                    else:
                        print(f"   ✅ {rel_path}/ (organized folder)")
                except:
                    print(f"   ✅ {rel_path}/ (organized folder)")
            if len(self.operations['organized']) > 5:
                print(f"   ... and {len(self.operations['organized']) - 5} more organized folders")
                
    def create_backup_copy(self):
        """Create backup of current state"""
        if not self.create_backup:
            return True
            
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        backup_dir = Path(f"./backup_{timestamp}")
        
        print(f"💾 Creating backup in {backup_dir}/...")
        
        try:
            backup_dir.mkdir(exist_ok=True)
            root = Path(".")
            
            # Only backup the Mice_* and PV_Mice* directories
            for mice_dir in root.iterdir():
                if mice_dir.is_dir() and (mice_dir.name.startswith("Mice_") or mice_dir.name.startswith("PV_Mice")):
                    copytree(mice_dir, backup_dir / mice_dir.name)
                    
            print(f"✅ Backup created successfully!")
            return backup_dir
            
        except Exception as e:
            print(f"❌ Backup failed: {e}")
            return False
            
    def get_user_confirmation(self):
        """Get user permission to proceed"""
        print(f"\n⚠️  Safety Check:")
        print("   - ✅ Write permissions confirmed")
        print("   - ✅ Preview generated successfully")
        
        if self.dry_run:
            print("\n🔍 DRY RUN MODE - No changes will be made")
            return True
            
        if self.create_backup:
            backup_choice = input("\n💾 Create backup before proceeding? (Y/n): ").strip().lower()
            self.create_backup = backup_choice != 'n'
            
        proceed = input("\n▶️  Proceed with organization? (y/N): ").strip().lower()
        return proceed == 'y'
        
    def perform_organization(self):
        """Execute the actual file operations"""
        if self.dry_run:
            print("\n🔍 DRY RUN - Would perform these operations:")
            return True
            
        print("\n🔄 Starting organization...")
        
        # Create backup if requested
        backup_location = None
        if self.create_backup:
            backup_location = self.create_backup_copy()
            if not backup_location:
                print("❌ Cannot proceed without backup. Aborting.")
                return False
                
        success_count = 0
        error_count = 0
        
        try:
            # Process moves
            current_dir = None
            for i, op in enumerate(self.operations['moves']):
                # Show progress for different directories
                op_dir = op['source'].parent
                if op_dir != current_dir:
                    current_dir = op_dir
                    rel_path = op_dir.relative_to(Path("."))
                    print(f"\n📁 Processing {rel_path}/...")
                    
                try:
                    # Create destination folder
                    op['dest_folder'].mkdir(exist_ok=True)
                    
                    # Double-check destination doesn't exist
                    if op['dest_file'].exists():
                        print(f"   ⚠️  Skipping {op['source'].name} - destination already exists")
                        error_count += 1
                        continue
                    
                    # Move file
                    move(op['source'], op['dest_file'])
                    
                    # Validate move was successful
                    if not op['dest_file'].exists():
                        raise Exception(f"Move failed: {op['dest_file']} not created")
                    if op['source'].exists():
                        raise Exception(f"Move failed: {op['source']} still exists")
                        
                    success_count += 1
                    
                    if self.verbose:
                        print(f"   📦 {op['source'].name} → {op['dest_folder'].name}/{op['dest_file'].name} ✅")
                        
                    # Show progress every 10 files
                    if (success_count + error_count) % 10 == 0:
                        print(f"   [Progress: {success_count + error_count}/{len(self.operations['moves'])} files]")
                        
                except Exception as e:
                    error_count += 1
                    print(f"   ❌ Failed to move {op['source'].name}: {e}")
                    continue
                    
            # Process deletions
            if self.operations['deletions']:
                print(f"\n🗑️ Cleaning up {len(self.operations['deletions'])} empty folders...")
                for folder in self.operations['deletions']:
                    try:
                        rmtree(folder)
                        if self.verbose:
                            rel_path = folder.relative_to(Path("."))
                            print(f"   🗑️ Deleted {rel_path} ✅")
                    except Exception as e:
                        print(f"   ❌ Failed to delete {folder}: {e}")
                        
            if error_count > 0:
                print(f"\n⚠️  Completed with {error_count} errors (see above)")
                
            return backup_location
            
        except Exception as e:
            print(f"❌ Critical error during organization: {e}")
            if backup_location:
                print(f"💡 Restore from backup: cp -r {backup_location}/* ./")
            return False
            
    def show_final_summary(self, backup_location=None):
        """Show completion summary and next steps"""
        print(f"\n✅ Organization Complete!")
        
        print(f"\n📊 SUMMARY:")
        print(f"   📦 {len(self.operations['moves'])} files moved into proper folders")
        print(f"   🗑️ {len(self.operations['deletions'])} empty folders deleted")
        print(f"   ✅ {len(self.operations['organized'])} files were already organized (left unchanged)")
        print(f"   📁 {self.stats['directories_processed']} mouse directories processed")
        
        if self.stats['skipped_exists'] > 0:
            print(f"   ⚠️  {self.stats['skipped_exists']} files skipped (already existed)")
        if self.stats['skipped_errors'] > 0:
            print(f"   ⚠️  {self.stats['skipped_errors']} items skipped due to errors")
            
        if backup_location and not self.dry_run:
            print(f"   📂 Backup saved to: {backup_location}/")
            
        print(f"\n🚀 NEXT STEPS:")
        print(f"   1. Verify the organization looks correct")
        print(f"   2. Run your main PNN pipeline script (e.g., batch_predict.py)")
        print(f"   3. Check logs for any processing issues")
        
        if backup_location and not self.dry_run:
            print(f"\n💡 TIP: If something went wrong, restore from backup:")
            print(f"   cp -r {backup_location}/* ./")

def show_help_guide():
    """Show detailed usage guide"""
    print("""
📖 PNN DATA ORGANIZATION GUIDE

PURPOSE:
This script organizes microscopy images for the PNN counting pipeline.
Each .tif image must be in its own folder for batch processing to work.

EXPECTED DIRECTORY STRUCTURE:
BEFORE:
  Mice_001/
    Mouse_01/
      WFA_1L.tif
      PV_2R.tif
      Control.tif

AFTER:
  Mice_001/
    Mouse_01/
      WFA_1L/
        WFA_1L.tif
      PV_2R/
        PV_2R.tif
      Control/
        Control.tif

USAGE:
  python structure.py              # Normal run with confirmations
  python structure.py --dry-run    # Preview only, no changes
  python structure.py --no-backup  # Skip backup creation
  python structure.py --verbose    # Show detailed progress

SAFETY FEATURES:
- Always shows preview before making changes
- Creates automatic backups (unless --no-backup)
- Only processes Mice_* and PV_Mice* folders
- Preserves existing organized folders
- Never deletes folders containing .csv or prediction files
- Skips files that already exist at destination
- Handles permission errors gracefully

💡 TIPS:
- Always run from the parent directory containing Mice_* folders
- Use --dry-run first to preview changes
- Keep backups of your original data structure
- Verify organization before running the main pipeline
""")

def main():
    parser = argparse.ArgumentParser(description='Organize microscopy data for PNN counting pipeline')
    parser.add_argument('--dry-run', action='store_true', 
                       help='Preview changes without executing them')
    parser.add_argument('--no-backup', action='store_true',
                       help='Skip backup creation')
    parser.add_argument('--verbose', action='store_true',
                       help='Show detailed progress information')
    parser.add_argument('--help-guide', action='store_true',
                       help='Show detailed usage guide and examples')
    
    args = parser.parse_args()
    
    if args.help_guide:
        show_help_guide()
        return
        
    # Initialize organizer
    organizer = DataOrganizer(
        dry_run=args.dry_run,
        create_backup=not args.no_backup,
        verbose=args.verbose
    )
    
    # Step 1: Check environment
    if not organizer.check_environment():
        sys.exit(1)
        
    # Step 2: Explain purpose
    organizer.explain_purpose()
    
    # Step 3: Scan and analyze
    organizer.scan_data_structure()
    
    # Step 4: Show analysis summary
    if not organizer.show_analysis_summary():
        print("\n🎉 Your data is ready for the PNN pipeline!")
        return
        
    # Step 5: Show detailed preview
    organizer.show_detailed_preview()
    
    # Step 6: Get user confirmation
    if not organizer.get_user_confirmation():
        print("\n❌ Operation cancelled by user.")
        return
        
    # Step 7: Perform organization
    backup_location = organizer.perform_organization()
    if backup_location is False:
        sys.exit(1)
        
    # Step 8: Show final summary
    organizer.show_final_summary(backup_location)

if __name__ == "__main__":
    main()