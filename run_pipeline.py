#!/usr/bin/env python3
"""
PNN/PV Cell Detection Pipeline Runner
Batch processes microscopy images through prediction and visualization pipeline
"""

import os
import sys
import subprocess
import argparse
import time
import json
import shutil
import re
from pathlib import Path
from datetime import datetime, timedelta
from collections import defaultdict

class PipelineRunner:
    def __init__(self, dry_run=False, verbose=True, skip_confirmation=False):
        self.dry_run = dry_run
        self.verbose = verbose
        self.skip_confirmation = skip_confirmation
        self.skip_partial = False  # Will be set by user prompt
        
        # Statistics tracking
        self.stats = {
            'total_files': 0,
            'already_processed': 0,
            'successfully_processed': 0,
            'failed': 0,
            'skipped': 0,
            'start_time': None,
            'files_to_process': []
        }
        
        # Error tracking
        self.errors = []
        
        # Time tracking for estimates
        self.processing_times = []
        
        # For progress animation
        self.spinner_chars = ['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏']
        self.spinner_index = 0
        
        # For tracking current subprocess info
        self.current_subprocess_info = {
            'dataset': '',
            'model': '',
            'device': '',
            'weights': '',
            'threshold': '',
            'output': ''
        }
        
    def check_environment(self):
        """Validate environment before running"""
        print("🔍 Checking environment...")
        
        # Check required scripts exist
        required_scripts = ['predict.py', 'draw_predictions.py']
        missing_scripts = []
        
        for script in required_scripts:
            if not Path(script).exists():
                missing_scripts.append(script)
                
        if missing_scripts:
            print(f"❌ Missing required scripts: {', '.join(missing_scripts)}")
            print("💡 Make sure you're running from the correct directory")
            return False
            
        # Check model folders exist
        model_folders = ['pnn_v2_fasterrcnn_640', 'pv_v2_fasterrcnn_640']
        missing_models = []
        
        for model in model_folders:
            if not Path(model).exists():
                missing_models.append(model)
                
        if missing_models:
            print(f"❌ Missing model folders: {', '.join(missing_models)}")
            print("💡 Download the required model files first")
            return False
            
        print("✅ All required files found")
        
        # Check disk space
        try:
            stat = shutil.disk_usage(".")
            free_gb = stat.free // (1024**3)
            if free_gb < 5:
                print(f"⚠️  Low disk space: {free_gb}GB free")
                print("💡 Each processed image may need 100-500MB")
            else:
                print(f"✅ Disk space OK: {free_gb}GB free")
        except:
            print("⚠️  Could not check disk space")
            
        # Check Python version
        if sys.version_info < (3, 6):
            print(f"❌ Python {sys.version} is too old. Need Python 3.6+")
            return False
            
        return True
        
    def explain_pipeline(self):
        """Explain what this pipeline does"""
        print("\n🎯 PNN/PV Cell Detection Pipeline")
        print("=" * 50)
        print("This pipeline processes microscopy images to:")
        print("  1. 🔬 Detect PNN or PV positive cells using AI")
        print("  2. 📊 Generate CSV files with cell locations")
        print("  3. 🖼️  Create annotated images showing detections")
        print("\nProcessing steps for each image:")
        print("  • Run AI prediction model (predict.py)")
        print("  • Generate visualization (draw_predictions.py)")
        print("=" * 50)
        
    def scan_for_work(self):
        """Scan directory structure and identify work to be done"""
        print("\n🔍 Scanning for images to process...")
        
        root = Path(".")
        work_items = []
        
        # Define which model to use for which folder
        folder_model_map = {
            'Mice_': 'pnn_v2_fasterrcnn_640',
            'PV_Mice': 'pv_v2_fasterrcnn_640'
        }
        
        for dir_pattern, model_name in folder_model_map.items():
            # Find matching directories
            if dir_pattern.endswith('_'):
                # Pattern like "Mice_*"
                matching_dirs = [
                    d for d in root.iterdir() 
                    if d.is_dir() and d.name.startswith(dir_pattern)
                ]
            else:
                # Exact match like "PV_Mice"
                matching_dirs = [
                    d for d in root.iterdir() 
                    if d.is_dir() and d.name == dir_pattern
                ]
                               
            for mice_dir in sorted(matching_dirs):
                if self.verbose:
                    print(f"  📁 Scanning {mice_dir.name}/...")
                    
                for mouse_dir in sorted(mice_dir.glob("Mouse_*")):
                    for subfolder in sorted(mouse_dir.iterdir()):
                        if not subfolder.is_dir():
                            continue
                            
                        # Check for .tif file
                        tif_files = (
                            list(subfolder.glob("*.tif")) + 
                            list(subfolder.glob("*.TIF"))
                        )
                        if not tif_files:
                            continue
                            
                        tif_file = tif_files[0]  # Use first .tif found
                        name = subfolder.name
                        csv_file = subfolder / f"localizations_{name}.csv"
                        preds_dir = subfolder / f"{name}_predictions"
                        
                        # Determine status
                        has_csv = csv_file.exists()
                        has_predictions = (
                            preds_dir.exists() and 
                            any(preds_dir.iterdir())
                        )
                        
                        status = 'pending'
                        if has_csv and has_predictions:
                            status = 'complete'
                            self.stats['already_processed'] += 1
                        elif has_csv or has_predictions:
                            status = 'partial'
                            
                        work_items.append({
                            'tif_file': tif_file,
                            'csv_file': csv_file,
                            'preds_dir': preds_dir,
                            'model': model_name,
                            'status': status,
                            'mice_dir': mice_dir.name,
                            'mouse_dir': mouse_dir.name,
                            'subfolder': subfolder.name
                        })
                        
        self.stats['total_files'] = len(work_items)
        
        return work_items
        
    def handle_partial_files_prompt(self, work_items):
        """Ask user what to do with partial files"""
        by_status = defaultdict(int)
        for item in work_items:
            by_status[item['status']] += 1
            
        if by_status['partial'] > 0 and not self.skip_confirmation:
            print(f"\n⚠️  Found {by_status['partial']} partially complete folders")
            print("   These have either CSV or images but not both.")
            print("\n   Examples of partial files:")
            
            # Show some examples
            partial_items = [w for w in work_items if w['status'] == 'partial']
            for item in partial_items[:3]:
                path = f"   • {item['mice_dir']}/{item['mouse_dir']}/{item['subfolder']}"
                if item['csv_file'].exists():
                    print(f"{path} (has CSV, missing images)")
                else:
                    print(f"{path} (has images, missing CSV)")
                    
            if len(partial_items) > 3:
                print(f"   ... and {len(partial_items) - 3} more")
                
            response = input("\n   Reprocess these partial files? (y/N): ")
            if response.strip().lower() != 'y':
                self.skip_partial = True
                print("   → Skipping partial files")
            else:
                print("   → Will reprocess partial files")
                
    def show_scan_summary(self, work_items):
        """Show summary of what was found"""
        print(f"\n📊 Scan Summary:")
        print(f"   Total images found: {len(work_items)}")
        
        by_status = defaultdict(int)
        for item in work_items:
            by_status[item['status']] += 1
            
        print(f"   ✅ Already complete: {by_status['complete']}")
        print(f"   ⏳ Ready to process: {by_status['pending']}")
        print(f"   ⚠️  Partially complete: {by_status['partial']}")
        
        # Filter files based on user choice
        if self.skip_partial:
            self.stats['files_to_process'] = [
                w for w in work_items 
                if w['status'] == 'pending'
            ]
        else:
            self.stats['files_to_process'] = [
                w for w in work_items 
                if w['status'] != 'complete'
            ]
            
        # Show examples of what will be processed
        if self.stats['files_to_process']:
            print(f"\n📋 Will process {len(self.stats['files_to_process'])} images:")
            for i, item in enumerate(self.stats['files_to_process'][:5]):
                path = f"{item['mice_dir']}/{item['mouse_dir']}/{item['subfolder']}"
                print(f"   • {path}")
            if len(self.stats['files_to_process']) > 5:
                remaining = len(self.stats['files_to_process']) - 5
                print(f"   ... and {remaining} more")
                
        # Time estimate
        if self.stats['files_to_process']:
            # Estimate 30 seconds per file
            total_files = len(self.stats['files_to_process'])
            estimated_time = total_files * 30
            print(f"\n⏱️  Estimated time: {self.format_time(estimated_time)}")
            print("   (Actual time depends on image size and computer speed)")
            
    def format_time(self, seconds):
        """Format seconds into human readable time"""
        if seconds < 60:
            return f"{int(seconds)} seconds"
        elif seconds < 3600:
            return f"{int(seconds/60)} minutes"
        else:
            hours = int(seconds/3600)
            minutes = int((seconds % 3600) / 60)
            return f"{hours}h {minutes}m"
    
    def create_progress_bar(self, current, total, width=30):
        """Create a progress bar string"""
        if total == 0:
            return "░" * width
        percent = current / total
        filled = int(width * percent)
        bar = "▓" * filled + "░" * (width - filled)
        return bar
    
    def clear_lines(self, num_lines):
        """Clear specified number of lines above cursor"""
        for _ in range(num_lines):
            print('\033[F\033[K', end='')
    
    def parse_subprocess_output(self, line):
        """Parse subprocess output to extract useful information"""
        # Parse predict.py output
        if '[  DATA]' in line:
            match = re.search(r'(\d+) image\(s\), (\d+) patches', line)
            if match:
                self.current_subprocess_info['dataset'] = f"{match.group(1)} image(s), {match.group(2)} patches"
        elif '[ MODEL]' in line:
            if 'FasterRCNN' in line:
                match = re.search(r'backbone=(\w+)', line)
                if match:
                    # Extract key model info
                    backbone = match.group(1)
                    nms_match = re.search(r'nms=([\d.]+)', line)
                    det_match = re.search(r'det_thresh=([\d.]+)', line)
                    nms = nms_match.group(1) if nms_match else '0.3'
                    det = det_match.group(1) if det_match else '0.05'
                    self.current_subprocess_info['model'] = f"FasterRCNN (backbone: {backbone}, NMS: {nms}, thresh: {det})"
        elif '[DEVICE]' in line:
            device = line.split('[DEVICE]')[1].strip()
            self.current_subprocess_info['device'] = device.upper()
        elif '[  CKPT]' in line:
            ckpt = line.split('[  CKPT]')[1].strip()
            self.current_subprocess_info['weights'] = ckpt
        elif '[PARAMS]' in line:
            match = re.search(r'thr = ([\d.]+)', line)
            if match:
                self.current_subprocess_info['threshold'] = match.group(1)
        elif '[OUTPUT]' in line:
            output = line.split('[OUTPUT]')[1].strip()
            # Shorten the path for display
            output_parts = output.split('\\')
            if len(output_parts) > 2:
                output = f"..\\{output_parts[-1]}"
            self.current_subprocess_info['output'] = output
    
    def display_overall_progress(self, current, total, elapsed_time, eta_time):
        """Display the overall progress box"""
        percent = (current / total * 100) if total > 0 else 0
        elapsed_str = self.format_time(elapsed_time)
        eta_str = self.format_time(eta_time) if eta_time > 0 else "calculating..."
        
        # Create the progress box
        progress_box = f"""╔═══════════════════════════════════════════════════════════════════════════════╗
║ 📊 Progress: {current}/{total} files ({percent:.1f}%) │ ⏱️  Elapsed: {elapsed_str:<10} │ 🚀 ETA: {eta_str:<10} ║
║ {self.create_progress_bar(current, total, width=67)} {percent:>3.0f}% ║
╚═══════════════════════════════════════════════════════════════════════════════╝"""
        
        print(progress_box)
        
    def display_detection_box(self, status_text="Initializing...", progress_percent=0):
        """Display the detection phase box"""
        progress_bar = self.create_progress_bar(progress_percent, 100, width=40)
        
        detection_box = f"""┌─ 🔬 Detection Phase ───────────────────────────────────────────────────────────┐
│ Dataset : {self.current_subprocess_info['dataset']:<67} │
│ Model   : {self.current_subprocess_info['model']:<67} │
│ Device  : {self.current_subprocess_info['device']:<67} │
│ Weights : {self.current_subprocess_info['weights']:<67} │
│ Config  : Detection threshold = {self.current_subprocess_info['threshold']:<50} │
│ Output  : {self.current_subprocess_info['output']:<67} │
│                                                                                 │
│ Status  : {progress_bar} {progress_percent:>3}% {status_text:<20} │
└─────────────────────────────────────────────────────────────────────────────────┘"""
        
        print(detection_box)
        
    def display_visualization_box(self, files_created, current_file=""):
        """Display the visualization phase box"""
        viz_box = f"""┌─ 🎨 Visualization Phase ────────────────────────────────────────────────────────┐
│ Creating overlay images...                                                      │"""
        
        # Show files being created
        for i, filename in enumerate(files_created[:4]):
            status = "✓" if i < len(files_created) - 1 or not current_file else "⠼ processing..."
            viz_box += f"\n│ • {filename:<63} {status:<10} │"
            
        if len(files_created) > 4:
            viz_box += f"\n│ • ... and {len(files_created) - 4} more                                                           │"
            
        viz_box += "\n└─────────────────────────────────────────────────────────────────────────────────┘"
        
        print(viz_box)
        
    def display_session_stats(self):
        """Display session statistics box"""
        success = self.stats['successfully_processed']
        failed = self.stats['failed']
        remaining = len(self.stats['files_to_process']) - success - failed
        
        avg_speed = "N/A"
        if self.processing_times and len(self.processing_times) > 0:
            avg_time = sum(self.processing_times) / len(self.processing_times)
            if avg_time > 0:
                files_per_min = 60 / avg_time
                avg_speed = f"{files_per_min:.1f}/m"
        
        stats_box = f"""╭─ Session Statistics ─╮
│ ✅ Success    : {success:<4} │
│ ❌ Failed     : {failed:<4} │
│ ⏭️  Remaining : {remaining:<4} │
│ 📈 Avg speed : {avg_speed:<5} │
╰──────────────────────╯"""
        
        print(stats_box)
    
    def run_subprocess_with_live_progress(self, cmd, work_item, overall_index, total_files):
        """Run subprocess with live progress updates"""
        
        # Start the process
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            universal_newlines=True,
            bufsize=1
        )
        
        output_lines = []
        is_detection = 'predict.py' in cmd
        is_visualization = 'draw_predictions.py' in cmd
        
        # Track created files for visualization
        created_files = []
        
        # Initial display
        if is_detection:
            # Clear subprocess info for new run
            self.current_subprocess_info = {
                'dataset': 'Loading...',
                'model': 'Loading...',
                'device': 'Detecting...',
                'weights': 'Loading...',
                'threshold': 'Loading...',
                'output': 'Configuring...'
            }
        
        # Read output line by line
        for line in process.stdout:
            output_lines.append(line)
            
            if is_detection:
                # Parse the line for information
                self.parse_subprocess_output(line)
                
                # Check for tqdm progress
                if 'PRED' in line and '%|' in line:
                    match = re.search(r'(\d+)%\|', line)
                    if match:
                        pct = int(match.group(1))
                        
                        # Clear previous display (11 lines for detection box)
                        self.clear_lines(11)
                        
                        # Redraw with updated progress
                        self.display_detection_box(f"Processing patches...", pct)
                        
                # Initial display once we have some info
                elif any(key in line for key in ['[  DATA]', '[ MODEL]', '[DEVICE]']):
                    self.clear_lines(11)
                    self.display_detection_box("Initializing...", 0)
                    
            elif is_visualization:
                # Check for "Saving:" messages
                if 'Saving:' in line:
                    filename = line.split('/')[-1].strip()
                    created_files.append(filename)
                    
                    # Clear and redraw visualization box
                    num_lines = min(len(created_files) + 3, 7)  # Box header + files + footer
                    self.clear_lines(num_lines)
                    self.display_visualization_box(created_files, filename)
        
        # Wait for process to complete
        process.wait()
        
        # Get the full output
        full_output = ''.join(output_lines)
        
        if process.returncode != 0:
            raise subprocess.CalledProcessError(
                process.returncode, cmd, output=full_output
            )
            
        return full_output, created_files
            
    def run_prediction_and_draw(self, work_item, overall_index, total_files):
        """Run prediction and drawing for a single image with live progress"""
        tif_file = work_item['tif_file']
        csv_file = work_item['csv_file']
        preds_dir = work_item['preds_dir']
        model = work_item['model']
        
        if self.dry_run:
            print(f"\n[DRY RUN] Would process {tif_file.name}")
            time.sleep(1)
            return True
            
        start_time = time.time()
        elapsed_total = time.time() - self.stats['start_time']
        
        # Calculate ETA
        eta_seconds = 0
        if self.processing_times:
            avg_time = sum(self.processing_times) / len(self.processing_times)
            remaining_files = total_files - overall_index
            eta_seconds = avg_time * remaining_files
        
        # Display overall progress
        print("\n" * 15)  # Make space for our display
        self.clear_lines(15)
        self.display_overall_progress(overall_index, total_files, elapsed_total, eta_seconds)
        
        print(f"\n📍 Processing: {work_item['mice_dir']} → {work_item['mouse_dir']} → {work_item['subfolder']}\n")
        
        # Step 1: Detection
        self.display_detection_box("Starting...", 0)
        
        # Run prediction
        cmd_predict = [
            sys.executable, "predict.py",
            model, str(tif_file),
            "-o", str(csv_file)
        ]
        
        try:
            output, _ = self.run_subprocess_with_live_progress(
                cmd_predict, work_item, overall_index, total_files
            )
            
            # Check if CSV was created
            if not csv_file.exists():
                raise Exception("CSV file was not created")
                
            # Count detections
            file_size_kb = csv_file.stat().st_size / 1024
            num_detections = "?"
            
            try:
                import pandas as pd
                df = pd.read_csv(csv_file)
                num_detections = len(df)
            except:
                # Can't count, but that's okay
                pass
            
            detection_time = time.time() - start_time
            print(f"\n✅ Detection complete: {num_detections} cells found ({file_size_kb:.1f} KB) in {detection_time:.1f}s\n")
            
        except subprocess.CalledProcessError as e:
            print(f"\n❌ Detection failed: {e}")
            self.errors.append(f"{tif_file}: Detection failed")
            return False
        except Exception as e:
            print(f"\n❌ Detection error: {e}")
            self.errors.append(f"{tif_file}: {e}")
            return False
            
        # Step 2: Visualization
        viz_start = time.time()
        
        # Create predictions directory
        preds_dir.mkdir(exist_ok=True)
        
        # Run drawing
        cmd_draw = [
            sys.executable, "draw_predictions.py",
            str(csv_file),
            "-r", str(tif_file.parent),
            "-o", str(preds_dir)
        ]
        
        try:
            output, created_files = self.run_subprocess_with_live_progress(
                cmd_draw, work_item, overall_index, total_files
            )
            
            # Check if images were created
            pred_images = list(preds_dir.glob("*.png")) + list(preds_dir.glob("*.jpg"))
            if not pred_images:
                raise Exception("No prediction images were created")
                
            viz_time = time.time() - viz_start
            total_time = time.time() - start_time
            
            # Final summary for this file
            print(f"\n✅ Visualization complete: {len(pred_images)} images created in {viz_time:.1f}s")
            print(f"⏱️  Total processing time: {total_time:.1f} seconds\n")
            
            # Display session stats
            self.display_session_stats()
            
        except subprocess.CalledProcessError as e:
            print(f"\n❌ Drawing failed: {e}")
            self.errors.append(f"{tif_file}: Drawing failed")
            return False
        except Exception as e:
            print(f"\n❌ Drawing error: {e}")
            self.errors.append(f"{tif_file}: {e}")
            return False
            
        # Track processing time
        self.processing_times.append(total_time)
        
        print("\n" + "═" * 85)
        
        return True
        
    def process_all_work(self, work_items):
        """Process all pending work items"""
        to_process = self.stats['files_to_process']
        
        if not to_process:
            return
            
        print(f"\n🚀 Starting batch processing...")
        print(f"   Processing {len(to_process)} images")
        print(f"   Press Ctrl+C to stop safely\n")
        
        self.stats['start_time'] = time.time()
        
        try:
            for i, work_item in enumerate(to_process):
                # Process the file with live progress
                success = self.run_prediction_and_draw(work_item, i + 1, len(to_process))
                
                if success:
                    self.stats['successfully_processed'] += 1
                else:
                    self.stats['failed'] += 1
                    
        except KeyboardInterrupt:
            print("\n\n⚠️  Processing interrupted by user")
            print("   Already completed files are saved")
            self.stats['skipped'] = len(to_process) - i - 1
            
    def get_user_confirmation(self):
        """Get user permission to proceed"""
        if self.skip_confirmation:
            return True
            
        if not self.stats['files_to_process']:
            print("\n✅ All images already processed! Nothing to do.")
            return False
            
        print("\n" + "="*50)
        
        if self.dry_run:
            print("🔍 DRY RUN MODE - No processing will occur")
            return True
            
        response = input(f"\n▶️  Ready to process {len(self.stats['files_to_process'])} images. Continue? (y/N): ")
        return response.strip().lower() == 'y'
        
    def save_processing_log(self):
        """Save a log of the processing session"""
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        log_file = Path(f"processing_log_{timestamp}.json")
        
        log_data = {
            'timestamp': timestamp,
            'stats': self.stats,
            'errors': self.errors,
            'parameters': {
                'dry_run': self.dry_run,
                'skip_partial': self.skip_partial
            }
        }
        
        try:
            with open(log_file, 'w') as f:
                json.dump(log_data, f, indent=2, default=str)
            print(f"\n📝 Processing log saved: {log_file}")
        except:
            pass
            
    def show_final_summary(self):
        """Show final processing summary"""
        print("\n" + "="*50)
        print("✅ Processing Complete!")
        print("="*50)
        
        if self.stats['start_time']:
            total_time = time.time() - self.stats['start_time']
            print(f"\n⏱️  Total time: {self.format_time(total_time)}")
            
        print(f"\n📊 Final Summary:")
        print(f"   Total images found: {self.stats['total_files']}")
        print(f"   Already processed: {self.stats['already_processed']}")
        print(f"   Successfully processed: {self.stats['successfully_processed']} ✅")
        print(f"   Failed: {self.stats['failed']} ❌")
        print(f"   Skipped: {self.stats['skipped']} ⏭️")
        
        if self.errors:
            print(f"\n❌ Errors encountered ({len(self.errors)}):")
            for error in self.errors[:5]:
                print(f"   • {error}")
            if len(self.errors) > 5:
                print(f"   ... and {len(self.errors) - 5} more")
            print("\n💡 Check the processing log for full error details")
            
        if self.stats['successfully_processed'] > 0:
            print(f"\n🎉 Successfully processed {self.stats['successfully_processed']} new images!")
            print("\n🚀 Next steps:")
            print("   1. Check the generated CSV files for cell counts")
            print("   2. Review prediction images in *_predictions folders")
            print("   3. Run any downstream analysis scripts")

def main():
    parser = argparse.ArgumentParser(
        description='Batch process microscopy images for cell detection',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python run_pipeline.py           # Normal run with default settings
  python run_pipeline.py --dry-run # Preview what would be processed
  python run_pipeline.py --quiet   # Less verbose output
  python run_pipeline.py -y        # Skip confirmation prompt
        """
    )
    
    parser.add_argument('--dry-run', action='store_true',
                       help='Preview what would be processed without running')
    parser.add_argument('--quiet', '-q', action='store_true',
                       help='Less verbose output')
    parser.add_argument('--yes', '-y', action='store_true',
                       help='Skip confirmation prompt')
    
    args = parser.parse_args()
    
    # Create runner
    runner = PipelineRunner(
        dry_run=args.dry_run,
        verbose=not args.quiet,
        skip_confirmation=args.yes
    )
    
    # Check environment
    if not runner.check_environment():
        sys.exit(1)
        
    # Explain what we're doing
    runner.explain_pipeline()
    
    # Scan for work
    work_items = runner.scan_for_work()
    
    # Handle partial files
    runner.handle_partial_files_prompt(work_items)
    
    # Show summary
    runner.show_scan_summary(work_items)
    
    # Get confirmation
    if not runner.get_user_confirmation():
        print("\n❌ Processing cancelled")
        return
        
    # Process all work
    runner.process_all_work(work_items)
    
    # Save log
    if not runner.dry_run:
        runner.save_processing_log()
        
    # Show final summary
    runner.show_final_summary()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        sys.exit(1)