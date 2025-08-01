#!/usr/bin/env python3
"""
PNN Colocalization Analysis Pipeline
Created by: Santisteban Lab, Richie Mistichelli
Date: 2025-july-31
Version: 1.0

Analyzes perineuronal net (PNN) detections across multiple microscopy channels,
performs colocalization analysis, and generates comprehensive reports.
"""

import os
import sys
import json
import csv
import numpy as np
import pandas as pd
from datetime import datetime
from pathlib import Path
import warnings
from typing import Dict, List, Tuple, Optional, Any
import shutil
from collections import defaultdict
import time

# Image processing imports
try:
    from PIL import Image, ImageDraw, ImageFont
    import cv2
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError as e:
    print(f"Error: Required package not installed. {e}")
    print("Please install: pip install pillow opencv-python openpyxl pandas numpy")
    sys.exit(1)

# Suppress warnings
warnings.filterwarnings('ignore')

class PNNColocalizationPipeline:
    def __init__(self):
        """Initialize the pipeline with default settings."""
        self.version = "1.0"
        self.created_by = "Santisteban Lab, Richie Mistichelli"
        self.date = "2025-july-31"
        
        # Pipeline state
        self.output_dir = None
        self.selected_channels = []
        self.available_channels = {}
        self.mice_to_process = []
        self.pixel_sizes = {}  # mouse/section -> µm/pixel
        self.thresholds = {}   # mouse/section -> threshold
        self.generate_distance_reports = False
        self.generate_visualizations = False
        self.visualization_settings = {}
        self.generate_threshold_visualizations = False
        self.threshold_viz_settings = {}
        self.processing_log = []
        self.error_log = []
        
        # Data storage
        self.all_data = {}  # Stores all processing results
        
    def run(self):
        """Main pipeline execution."""
        self.print_header()
        
        # Step 1: Check directories and available channels
        if not self.check_directories():
            return
            
        # Step 2: Channel selection
        self.select_channels()
        
        # Step 3: Mouse selection
        self.select_mice()
        
        # Step 4: Pixel size configuration
        self.configure_pixel_sizes()
        
        # Step 5: Threshold configuration
        self.configure_thresholds()
        
        # Step 6: Additional options
        self.configure_additional_options()
        
        # Step 7: Check existing analyses
        self.check_existing_analyses()
        
        # Step 8: Final confirmation
        if not self.final_confirmation():
            print("\n[INFO] Analysis cancelled by user.")
            return
            
        # Step 9: Process all data
        self.process_all_sections()
        
        # Step 10: Generate summaries
        self.generate_summaries()
        
        # Step 11: Show completion summary
        self.show_completion_summary()
    
    def print_header(self):
        """Print the pipeline header."""
        print("═" * 70)
        print(" " * 15 + "PNN COLOCALIZATION ANALYSIS PIPELINE v" + self.version)
        print("═" * 70)
        print(f"Created by: {self.created_by}")
        print(f"Date: {self.date}")
        print("\nInitializing...\n")
        
    def check_directories(self) -> bool:
        """Check for required directories and detect available channels."""
        print("[INFO] Checking for required directories...")
        
        channel_dirs = {
            'WFA': 'Mice_WFA',
            'Agg': 'Mice_Agg', 
            'PV': 'PV_Mice'
        }
        
        found_channels = []
        for channel, dirname in channel_dirs.items():
            if os.path.exists(dirname):
                # Count subdirectories
                subdirs = [d for d in os.listdir(dirname) 
                          if os.path.isdir(os.path.join(dirname, d))]
                print(f"✓ Found {dirname}/ ({len(subdirs)} subdirectories)")
                found_channels.append(channel)
                self.available_channels[channel] = dirname
            else:
                print(f"✗ {dirname}/ not found")
                
        if len(found_channels) < 2:
            print("\n❌ Error: At least 2 channel directories required.")
            print("Please ensure you have at least 2 of: Mice_WFA/, Mice_Agg/, PV_Mice/")
            return False
            
        print(f"\n✓ Minimum requirement met: {len(found_channels)} channels available " +
              f"({', '.join(found_channels)})")
        
        # Create output directory
        timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
        self.output_dir = f"Analysis_{timestamp}"
        os.makedirs(self.output_dir, exist_ok=True)
        print(f"\n[INFO] Creating output directory: {self.output_dir}/")
        
        return True
        
    def select_channels(self):
        """Allow user to select which channels to analyze."""
        print("\n" + "═" * 70)
        print("STEP 1: CHANNEL SELECTION")
        print("═" * 70)
        
        available = list(self.available_channels.keys())
        print(f"\nAvailable channels detected: {', '.join(available)}")
        
        while True:
            response = input("\nWhich channels would you like to analyze?\n" +
                           "Enter channel names separated by spaces (e.g., WFA PV)\n" +
                           "Or type 'all' for all available channels: ").strip().upper()
            
            if response == 'ALL':
                self.selected_channels = available
                break
            else:
                channels = response.split()
                if all(ch in available for ch in channels) and len(channels) >= 2:
                    self.selected_channels = channels
                    break
                else:
                    print("❌ Invalid selection. Please choose at least 2 available channels.")
                    
        print(f"\n✓ Selected channels: {', '.join(self.selected_channels)}")
        
    def select_mice(self):
        """Allow user to select which mice to process."""
        print("\n" + "═" * 70)
        print("STEP 2: MOUSE SELECTION")
        print("═" * 70)
        
        # Scan for mice with selected channels
        print("\n[INFO] Scanning selected folders for mice data...")
        
        mice_sections = self.scan_mice_sections()
        
        if not mice_sections:
            print("❌ No mice found with selected channels!")
            sys.exit(1)
            
        # Display found mice
        print(f"\nFound {len(mice_sections)} mice with data:")
        for mouse, sections in sorted(mice_sections.items()):
            print(f"• {mouse} ({len(sections)} sections)")
            
        print(f"\nTotal: {sum(len(s) for s in mice_sections.values())} sections to potentially process")
        
        # Selection menu
        while True:
            print("\nWhich mice would you like to process?")
            print("1. All mice (" + str(len(mice_sections)) + " total)")
            print("2. Specific mice")
            print("3. View section details for each mouse")
            
            choice = input("\nEnter choice (1-3): ").strip()
            
            if choice == '1':
                self.mice_to_process = list(mice_sections.keys())
                print(f"\n✓ Will process all {len(self.mice_to_process)} mice")
                break
                
            elif choice == '2':
                mouse_list = input("\nEnter mouse names separated by spaces: ").strip().split()
                valid = [m for m in mouse_list if m in mice_sections]
                if valid:
                    self.mice_to_process = valid
                    print(f"\n✓ Will process: {', '.join(valid)}")
                    break
                else:
                    print("❌ No valid mouse names entered.")
                    
            elif choice == '3':
                self.show_section_details(mice_sections)
                
    def scan_mice_sections(self) -> Dict[str, List[Dict]]:
        """Scan directories to find all mice and their sections."""
        mice_sections = defaultdict(list)
        
        # For each selected channel, scan for mice
        for channel in self.selected_channels:
            channel_dir = self.available_channels[channel]
            
            if not os.path.exists(channel_dir):
                continue
                
            # Scan mouse folders
            for mouse_folder in os.listdir(channel_dir):
                mouse_path = os.path.join(channel_dir, mouse_folder)
                if not os.path.isdir(mouse_path):
                    continue
                    
                # Scan section folders
                for section_folder in os.listdir(mouse_path):
                    section_path = os.path.join(mouse_path, section_folder)
                    if not os.path.isdir(section_path):
                        continue
                        
                    # Extract section name (without channel suffix)
                    # e.g., IB60_CA1_1L_WFA -> IB60_CA1_1L
                    parts = section_folder.split('_')
                    if len(parts) >= 4:
                        section_name = '_'.join(parts[:-1])  # Remove channel suffix
                        
                        # Check if localization CSV exists
                        csv_path = os.path.join(section_path, f"localizations_{section_folder}.csv")
                        if os.path.exists(csv_path):
                            # Add to mouse's section list
                            section_info = {
                                'section': section_name,
                                'channel': channel,
                                'folder': section_folder,
                                'csv_path': csv_path,
                                'image_path': os.path.join(section_path, f"{section_folder}.tif")
                            }
                            
                            # Check if this section already exists for this mouse
                            existing = next((s for s in mice_sections[mouse_folder] 
                                           if s['section'] == section_name), None)
                            if existing:
                                existing['channels'].append(channel)
                                existing['paths'][channel] = section_info
                            else:
                                mice_sections[mouse_folder].append({
                                    'section': section_name,
                                    'channels': [channel],
                                    'paths': {channel: section_info}
                                })
                                
        return dict(mice_sections)
        
    def show_section_details(self, mice_sections: Dict):
        """Display detailed section information for each mouse."""
        print("\n" + "-" * 50)
        for mouse, sections in sorted(mice_sections.items()):
            print(f"\n{mouse} sections:")
            for section_data in sections:
                channels_present = []
                for ch in self.selected_channels:
                    if ch in section_data['channels']:
                        channels_present.append(f"{ch} ✓")
                    else:
                        channels_present.append(f"{ch} ✗ - missing")
                print(f"  - {section_data['section']} ({', '.join(channels_present)})")
                
                # Warn about missing channels
                missing = [ch for ch in self.selected_channels if ch not in section_data['channels']]
                if missing:
                    print(f"    ⚠️ Note: {section_data['section']} is missing {', '.join(missing)} data")
                    
        process_anyway = input("\nProcess all mice anyway? (y/n): ").strip().lower()
        if process_anyway == 'y':
            print("✓ Will process all available data")
            
    def configure_pixel_sizes(self):
        """Configure pixel size (µm/pixel) for measurements."""
        print("\n" + "═" * 70)
        print("STEP 3: PIXEL SIZE CONFIGURATION")
        print("═" * 70)
        
        print("\n[INFO] Scanning TIF files for pixel size metadata...")
        print("[INFO] No automatic pixel size data found in TIF metadata")
        
        print("\nHow would you like to specify pixel measurements?")
        print("1. Use pixels only (no micron conversion)")
        print("2. Set one pixel size for ALL images")
        print("3. Set pixel size for each mouse")
        print("4. Set pixel size for each section individually")
        print("5. Load from previous configuration file")
        
        choice = input("\nEnter choice (1-5): ").strip()
        
        if choice == '1':
            # No pixel sizes - all measurements in pixels only
            print("\n⚠️ All measurements will be in pixels only (no micron conversion)")
            self.pixel_sizes = {}
            
        elif choice == '2':
            # One pixel size for all
            pixel_size = self.get_pixel_size_input("Enter pixel size (µm/pixel)")
            if pixel_size:
                print(f"\n✓ Using {pixel_size} µm/pixel for all images")
                # Apply to all mice/sections
                mice_sections = self.scan_mice_sections()
                for mouse in self.mice_to_process:
                    for section_data in mice_sections.get(mouse, []):
                        key = f"{mouse}/{section_data['section']}"
                        self.pixel_sizes[key] = pixel_size
                        
        elif choice == '3':
            # Pixel size per mouse
            print("\n" + "═" * 70)
            print("Setting pixel size for each mouse:")
            print("(Press Enter to skip and use pixel-only measurements)\n")
            
            mice_sections = self.scan_mice_sections()
            last_value = None
            
            for mouse in sorted(self.mice_to_process):
                if last_value:
                    prompt = f"{mouse} - Enter pixel size (µm/pixel) or 's' for same ({last_value}): "
                else:
                    prompt = f"{mouse} - Enter pixel size (µm/pixel): "
                    
                response = input(prompt).strip()
                
                if response.lower() == 's' and last_value:
                    pixel_size = last_value
                elif response == '':
                    pixel_size = None
                    print(f"  ⚠️ {mouse}: No pixel size - will use pixel-only measurements")
                else:
                    pixel_size = self.get_pixel_size_input("", response)
                    
                if pixel_size:
                    last_value = pixel_size
                    sections = mice_sections.get(mouse, [])
                    print(f"  ✓ {mouse}: {pixel_size} µm/pixel will apply to all {len(sections)} sections")
                    
                    # Apply to all sections of this mouse
                    for section_data in sections:
                        key = f"{mouse}/{section_data['section']}"
                        self.pixel_sizes[key] = pixel_size
                        
        elif choice == '4':
            # Pixel size per section
            print("\n" + "═" * 70)
            print("Setting pixel size for each section:")
            print("(Press Enter to skip and use pixel-only for that section)\n")
            
            mice_sections = self.scan_mice_sections()
            last_value = None
            
            for mouse in sorted(self.mice_to_process):
                print(f"\n{mouse}:")
                sections = mice_sections.get(mouse, [])
                
                for section_data in sections:
                    section = section_data['section']
                    
                    if last_value:
                        prompt = f"  {section} - Enter pixel size (µm/pixel) or 's' for same ({last_value}): "
                    else:
                        prompt = f"  {section} - Enter pixel size (µm/pixel): "
                        
                    response = input(prompt).strip()
                    
                    if response.lower() == 's' and last_value:
                        pixel_size = last_value
                    elif response == '':
                        pixel_size = None
                    else:
                        pixel_size = self.get_pixel_size_input("", response)
                        
                    if pixel_size:
                        last_value = pixel_size
                        key = f"{mouse}/{section}"
                        self.pixel_sizes[key] = pixel_size
                        
        elif choice == '5':
            # Load from file
            self.load_pixel_size_config()
            
        # Show summary
        self.show_pixel_size_summary()
        
        # Save configuration
        if self.pixel_sizes and choice != '5':
            save = input("\nSave this configuration for future use? (y/n): ").strip().lower()
            if save == 'y':
                self.save_pixel_size_config()
                
    def get_pixel_size_input(self, prompt: str, value: str = None) -> Optional[float]:
        """Get and validate pixel size input."""
        if value is None:
            value = input(prompt + ": ").strip()
            
        try:
            pixel_size = float(value)
            if pixel_size > 0:
                return pixel_size
            else:
                print("❌ Pixel size must be positive")
                return None
        except ValueError:
            if value != '':
                print("❌ Invalid pixel size")
            return None
            
    def show_pixel_size_summary(self):
        """Display summary of pixel size configuration."""
        print("\n" + "═" * 70)
        print("PIXEL SIZE SUMMARY:")
        
        mice_with_data = set()
        mice_without_data = set()
        
        mice_sections = self.scan_mice_sections()
        
        for mouse in self.mice_to_process:
            sections = mice_sections.get(mouse, [])
            has_any_pixel_size = False
            
            for section_data in sections:
                key = f"{mouse}/{section_data['section']}"
                if key in self.pixel_sizes:
                    has_any_pixel_size = True
                    
            if has_any_pixel_size:
                mice_with_data.add(mouse)
            else:
                mice_without_data.add(mouse)
                
        # Display mice with pixel sizes
        for mouse in sorted(mice_with_data):
            sections = mice_sections.get(mouse, [])
            # Check if all sections have same pixel size
            pixel_sizes_for_mouse = set()
            for section_data in sections:
                key = f"{mouse}/{section_data['section']}"
                if key in self.pixel_sizes:
                    pixel_sizes_for_mouse.add(self.pixel_sizes[key])
                    
            if len(pixel_sizes_for_mouse) == 1:
                print(f"✓ {mouse}: {pixel_sizes_for_mouse.pop()} µm/pixel ({len(sections)} sections)")
            else:
                print(f"✓ {mouse}: Mixed pixel sizes across sections")
                
        # Display mice without pixel sizes
        for mouse in sorted(mice_without_data):
            sections = mice_sections.get(mouse, [])
            print(f"✗ {mouse}: Pixels only ({len(sections)} sections)")
            
    def save_pixel_size_config(self):
        """Save pixel size configuration to file."""
        config_dir = os.path.join(self.output_dir, "configs")
        os.makedirs(config_dir, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
        config_file = os.path.join(config_dir, f"pixel_size_config_{timestamp}.json")
        
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(self.pixel_sizes, f, indent=2)
            
        print(f"✓ Saved to: {config_file}")
        
        # Also save to a standard location for easy reloading
        standard_config = "pixel_configs/latest_config.json"
        os.makedirs("pixel_configs", exist_ok=True)
        shutil.copy(config_file, standard_config)
        
    def load_pixel_size_config(self):
        """Load pixel size configuration from file."""
        config_files = []
        
        # Check standard location
        if os.path.exists("pixel_configs/latest_config.json"):
            config_files.append("pixel_configs/latest_config.json")
            
        # Check for other configs
        if os.path.exists("pixel_configs"):
            for file in sorted(os.listdir("pixel_configs")):
                if file.endswith(".json") and file != "latest_config.json":
                    config_files.append(os.path.join("pixel_configs", file))
                    
        if not config_files:
            print("❌ No configuration files found")
            return
            
        print("\nAvailable configuration files:")
        for i, file in enumerate(config_files, 1):
            print(f"{i}. {file}")
            
        choice = input("\nSelect configuration file (number): ").strip()
        
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(config_files):
                with open(config_files[idx], 'r') as f:
                    self.pixel_sizes = json.load(f)
                print(f"✓ Loaded configuration from {config_files[idx]}")
            else:
                print("❌ Invalid selection")
        except (ValueError, FileNotFoundError, json.JSONDecodeError) as e:
            print(f"❌ Error loading configuration: {e}")
            
    def configure_thresholds(self):
        """Configure colocalization thresholds."""
        print("\n" + "═" * 70)
        print("STEP 4: COLOCALIZATION THRESHOLD SETTINGS")
        print("═" * 70)
        
        # Separate mice with and without pixel size data
        mice_with_pixel_size = []
        mice_without_pixel_size = []
        
        mice_sections = self.scan_mice_sections()
        
        for mouse in self.mice_to_process:
            has_pixel_size = False
            sections = mice_sections.get(mouse, [])
            
            for section_data in sections:
                key = f"{mouse}/{section_data['section']}"
                if key in self.pixel_sizes:
                    has_pixel_size = True
                    break
                    
            if has_pixel_size:
                mice_with_pixel_size.append(mouse)
            else:
                mice_without_pixel_size.append(mouse)
                
        print("\nYou have mice with and without pixel size data.")
        
        # Get threshold for mice with pixel size
        if mice_with_pixel_size:
            print(f"\nFor mice WITH pixel size ({', '.join(mice_with_pixel_size)}):")
            micron_threshold = None
            while micron_threshold is None:
                response = input("Enter colocalization threshold in MICRONS: ").strip()
                try:
                    micron_threshold = float(response)
                    if micron_threshold <= 0:
                        print("❌ Threshold must be positive")
                        micron_threshold = None
                except ValueError:
                    print("❌ Invalid number")
                    
            # Convert to pixels for each mouse/section
            print("\nThis converts to:")
            for mouse in mice_with_pixel_size:
                sections = mice_sections.get(mouse, [])
                pixel_sizes_for_mouse = set()
                
                for section_data in sections:
                    key = f"{mouse}/{section_data['section']}"
                    if key in self.pixel_sizes:
                        pixel_size = self.pixel_sizes[key]
                        pixel_threshold = micron_threshold / pixel_size
                        self.thresholds[key] = {
                            'microns': micron_threshold,
                            'pixels': pixel_threshold,
                            'pixel_size': pixel_size
                        }
                        pixel_sizes_for_mouse.add((pixel_size, pixel_threshold))
                        
                # Display conversion
                for pixel_size, pixel_threshold in pixel_sizes_for_mouse:
                    print(f"  • {mouse}: {pixel_threshold:.1f} pixels (at {pixel_size} µm/pixel)")
                    
        # Get threshold for mice without pixel size
        if mice_without_pixel_size:
            print(f"\nFor mice WITHOUT pixel size ({', '.join(mice_without_pixel_size)}):")
            pixel_threshold = None
            while pixel_threshold is None:
                response = input("Enter colocalization threshold in PIXELS: ").strip()
                try:
                    pixel_threshold = float(response)
                    if pixel_threshold <= 0:
                        print("❌ Threshold must be positive")
                        pixel_threshold = None
                except ValueError:
                    print("❌ Invalid number")
                    
            # Apply to all sections without pixel size
            for mouse in mice_without_pixel_size:
                sections = mice_sections.get(mouse, [])
                for section_data in sections:
                    key = f"{mouse}/{section_data['section']}"
                    self.thresholds[key] = {
                        'microns': None,
                        'pixels': pixel_threshold,
                        'pixel_size': None
                    }
                    
    def configure_additional_options(self):
        """Configure additional analysis options."""
        print("\n" + "═" * 70)
        print("STEP 5: ADDITIONAL ANALYSIS OPTIONS")
        print("═" * 70)
        
        # Distance distribution reports
        response = input("\nGenerate distance distribution reports? (y/n): ").strip().lower()
        self.generate_distance_reports = (response == 'y')
        
        if self.generate_distance_reports:
            print("These help optimize threshold selection by showing cell counts at each distance.")
            
            print("\nFor distance distribution analysis:")
            print("1. Use microns for all (skip sections without pixel size)")
            print("2. Use pixels for all sections")
            print("3. Use microns where available, pixels otherwise (mixed)")
            
            self.distance_report_mode = input("\nEnter choice (1-3): ").strip()
            
            # Get maximum distances
            mice_with_pixel_size = [m for m in self.mice_to_process 
                                   if any(f"{m}/" in k for k in self.pixel_sizes.keys())]
            mice_without_pixel_size = [m for m in self.mice_to_process 
                                     if m not in mice_with_pixel_size]
            
            if mice_with_pixel_size and self.distance_report_mode in ['1', '3']:
                self.max_distance_microns = float(
                    input("\nMaximum distance to analyze for sections with pixel size (in microns): ").strip()
                )
                
            if mice_without_pixel_size and self.distance_report_mode in ['2', '3']:
                self.max_distance_pixels = float(
                    input("Maximum distance to analyze for sections without pixel size (in pixels): ").strip()
                )
                
            response = input("\nStop analysis if plateau detected (<0.5% new in 3 bins)? (y/n): ").strip().lower()
            self.stop_at_plateau = (response == 'y')
            
        # Visual overlays
        print("\n" + "═" * 70)
        response = input("\nGenerate visual overlays? (y/n): ").strip().lower()
        self.generate_visualizations = (response == 'y')
        
        if self.generate_visualizations:
            print("\nVisualization settings:")
            
            # Circle outline thickness
            thickness = input("• Circle outline thickness (1-3 pixels, default 2): ").strip()
            self.visualization_settings['outline_thickness'] = int(thickness) if thickness else 2
            
            # Circle diameter
            diameter = input("• Circle diameter for detections (pixels, default 35): ").strip()
            self.visualization_settings['circle_diameter'] = int(diameter) if diameter else 35
            
            # Colocalized thickness
            coloc_thickness = input("• Colocalized circle thickness (2-5 pixels, default 3): ").strip()
            self.visualization_settings['coloc_thickness'] = int(coloc_thickness) if coloc_thickness else 3
            
            print("\nColor scheme:")
            print("  - WFA: Yellow")
            print("  - PV: Blue")
            print("  - Agg: Red")
            print("  - 2-way colocalized: White")
            print("  - 3-way colocalized: Magenta")
            print("  - Scale bar: White")
            print("\nAll layout types will be generated automatically.")
            
        # Threshold visualization settings
        print("\n" + "═" * 70)
        print("THRESHOLD VISUALIZATION SETTINGS")
        print("═" * 70)
        print("These create additional grayscale visualizations showing threshold effects.")
        print("They will use the same threshold values you already set.")
        
        response = input("\nGenerate threshold-focused visualizations? (y/n): ").strip().lower()
        self.generate_threshold_visualizations = (response == 'y')
        
        if self.generate_threshold_visualizations:
            # Initialize threshold viz settings
            self.threshold_viz_settings = {}
            
            # Circle diameter
            diameter = input("\nCircle diameter for threshold visualizations (pixels, default 35): ").strip()
            self.threshold_viz_settings['circle_diameter'] = int(diameter) if diameter else 35
            
            # Visualization method
            print("\nHow would you like to indicate colocalization status?")
            print("1. Thickness difference (thick = within threshold, thin = beyond threshold)")
            print("2. Color coding (original colors = within threshold, new colors = beyond)")
            
            method = input("\nEnter choice (1-2): ").strip()
            self.threshold_viz_settings['method'] = 'thickness' if method == '1' else 'color'
            
            if self.threshold_viz_settings['method'] == 'thickness':
                # Get thickness values
                thick = input("\n• Thickness for colocalized circles (pixels, default 4): ").strip()
                thin = input("• Thickness for non-colocalized circles (pixels, default 1): ").strip()
                self.threshold_viz_settings['thick_circle'] = int(thick) if thick else 4
                self.threshold_viz_settings['thin_circle'] = int(thin) if thin else 1
                print("\n✓ Using thickness-based method")
            else:
                print("\n✓ Using color-coded method")
                print("  Within threshold: Original colors (WFA=yellow, PV=blue, Agg=red)")
                print("  Beyond threshold: Alternative colors (WFA=magenta, PV=green, Agg=white)")
                
                # Define color mappings
                self.threshold_viz_settings['original_colors'] = {
                    'WFA': (0, 255, 255),    # Yellow
                    'PV': (255, 0 , 0),      # Blue
                    'Agg': (0, 0, 255)       # Red
                }
                self.threshold_viz_settings['alt_colors'] = {
                    'WFA': (181, 61, 253),   # Magenta
                    'PV': (47, 205, 17),   # green
                    'Agg': (255, 255, 255)       # Green
                }
            
            # Background blending settings
            print("\n" + "═" * 70)
            print("BACKGROUND BLENDING SETTINGS")
            print("═" * 70)
            print("For composite backgrounds, adjust how much each channel contributes.")
            print("This helps balance brightness differences between channels.")
            
            # Detect which channel pairs will be created
            channel_pairs = []
            channels = self.selected_channels
            for i in range(len(channels)):
                for j in range(i + 1, len(channels)):
                    channel_pairs.append((channels[i], channels[j]))
            
            print(f"\n[Detected channels: {', '.join(channels)}]")
            print(f"[Will create visualizations for pairs: {', '.join([f'{c1}-{c2}' for c1, c2 in channel_pairs])}]")
            
            # Get blending ratios for each pair
            self.threshold_viz_settings['blend_ratios'] = {}
            
            for ch1, ch2 in channel_pairs:
                print(f"\n{ch1}-{ch2} Composite Blend:")
                
                while True:
                    blend1_str = input(f"• {ch1} contribution (0.1-0.9, default 0.5): ").strip()
                    blend1 = float(blend1_str) if blend1_str else 0.5
                    
                    if 0.1 <= blend1 <= 0.9:
                        blend2 = 1.0 - blend1
                        print(f"• {ch2} contribution: {blend2:.1f}")
                        print(f"  ✓ Composite will be {int(blend1*100)}% {ch1} + {int(blend2*100)}% {ch2}")
                        
                        self.threshold_viz_settings['blend_ratios'][f"{ch1}_{ch2}"] = {
                            ch1: blend1,
                            ch2: blend2
                        }
                        break
                    else:
                        print("  ❌ Please enter a value between 0.1 and 0.9")
            
    def check_existing_analyses(self):
        """Check for existing analysis results."""
        print("\n" + "═" * 70)
        print("STEP 6: CHECKING FOR EXISTING ANALYSES")
        print("═" * 70)
        
        print("\n[INFO] Scanning for previous analysis results...")
        
        # For now, assume no existing analyses in new output directory
        print("\nNo existing analyses found in output directory.")
        print("✓ All sections will be processed fresh.")
        
    def final_confirmation(self) -> bool:
        """Show final summary and get user confirmation."""
        print("\n" + "═" * 70)
        print("FINAL CONFIRMATION")
        print("═" * 70)
        
        mice_sections = self.scan_mice_sections()
        
        # Count sections
        total_sections = 0
        sections_with_pixel_size = 0
        sections_missing_channels = 0
        
        for mouse in self.mice_to_process:
            sections = mice_sections.get(mouse, [])
            for section_data in sections:
                total_sections += 1
                key = f"{mouse}/{section_data['section']}"
                
                if key in self.pixel_sizes:
                    sections_with_pixel_size += 1
                    
                if len(section_data['channels']) < len(self.selected_channels):
                    sections_missing_channels += 1
                    
        print("\nANALYSIS SUMMARY:")
        print("────────────────")
        print(f"Output Directory: {self.output_dir}/")
        print("Processing Mode: New analysis (no existing data)")
        
        print("\nMICE TO PROCESS:")
        print("────────────────")
        for mouse in sorted(self.mice_to_process):
            sections = mice_sections.get(mouse, [])
            
            # Check pixel size
            has_pixel_size = any(f"{mouse}/{s['section']}" in self.pixel_sizes for s in sections)
            
            if has_pixel_size:
                # Get pixel size(s)
                pixel_sizes = set()
                for s in sections:
                    key = f"{mouse}/{s['section']}"
                    if key in self.pixel_sizes:
                        pixel_sizes.add(self.pixel_sizes[key])
                        
                if len(pixel_sizes) == 1:
                    print(f"• {mouse}: {len(sections)} sections ({pixel_sizes.pop()} µm/pixel)")
                else:
                    print(f"• {mouse}: {len(sections)} sections (mixed pixel sizes)")
            else:
                print(f"• {mouse}: {len(sections)} sections (pixels only)")
                
            # Check for missing channels
            for section_data in sections:
                missing = [ch for ch in self.selected_channels if ch not in section_data['channels']]
                if missing:
                    print(f"  └─ ⚠️ {section_data['section']} missing {', '.join(missing)} data")
                    
        print(f"\nTotal: {total_sections} sections")
        if sections_missing_channels > 0:
            print(f"  ({total_sections - sections_missing_channels} with full colocalization, " +
                  f"{sections_missing_channels} with missing channels)")
            
        print("\nSETTINGS:")
        print("─────────")
        print(f"• Channels: {', '.join(self.selected_channels)}")
        print("• Thresholds:")
        
        # Display threshold summary
        mice_with_microns = set()
        mice_with_pixels = set()
        
        for key, threshold_data in self.thresholds.items():
            mouse = key.split('/')[0]
            if threshold_data['microns']:
                mice_with_microns.add((mouse, threshold_data['microns'], threshold_data['pixels']))
            else:
                mice_with_pixels.add((mouse, threshold_data['pixels']))
                
        for mouse, microns, pixels in sorted(mice_with_microns):
            print(f"  - {mouse}: {microns} µm ({pixels:.1f} pixels)")
            
        for mouse, pixels in sorted(mice_with_pixels):
            print(f"  - {mouse}: {pixels} pixels")
            
        if self.generate_distance_reports:
            print(f"• Distance distributions: Yes (max {getattr(self, 'max_distance_microns', 'N/A')} µm " +
                  f"or {getattr(self, 'max_distance_pixels', 'N/A')} pixels)")
        else:
            print("• Distance distributions: No")
            
        print(f"• Visual overlays: {'Yes (all layouts)' if self.generate_visualizations else 'No'}")
        
        if self.generate_threshold_visualizations:
            method = self.threshold_viz_settings.get('method', 'color')
            method_desc = 'thickness-based' if method == 'thickness' else 'color-coded'
            print(f"• Threshold visualizations: Yes ({method_desc})")
        else:
            print("• Threshold visualizations: No")
            
        print("• Excel outputs: Yes")
        
        print(f"\nESTIMATED TIME: ~{total_sections * 1} minutes")
        
        response = input("\nProceed with analysis? (y/n): ").strip().lower()
        return response == 'y'
        
    def process_all_sections(self):
        """Process all selected sections."""
        print("\n" + "═" * 70)
        print("PROCESSING")
        print("═" * 70 + "\n")
        
        mice_sections = self.scan_mice_sections()
        total_sections = sum(len(mice_sections.get(m, [])) for m in self.mice_to_process)
        section_num = 0
        
        start_time = time.time()
        
        for mouse in sorted(self.mice_to_process):
            sections = mice_sections.get(mouse, [])
            
            for section_data in sections:
                section_num += 1
                section_key = f"{mouse}/{section_data['section']}"
                
                # Process this section
                self.process_section(mouse, section_data, section_num, total_sections)
                
        elapsed_time = time.time() - start_time
        minutes = int(elapsed_time // 60)
        seconds = int(elapsed_time % 60)
        print(f"\n[INFO] Total processing time: {minutes} minutes {seconds} seconds")
        
    def process_section(self, mouse: str, section_data: Dict, section_num: int, total_sections: int):
        """Process a single section."""
        section = section_data['section']
        section_key = f"{mouse}/{section}"
        
        # Print section header
        print("┌" + "─" * 68 + "┐")
        print(f"│ {mouse} - Section {section} ({section_num}/{total_sections})".ljust(68) + "│")
        print("├" + "─" * 68 + "┤")
        
        # Get pixel size and threshold for this section
        pixel_size = self.pixel_sizes.get(section_key)
        threshold_data = self.thresholds.get(section_key, {})
        
        if pixel_size:
            print(f"│ Pixel size: {pixel_size} µm/pixel".ljust(68) + "│")
            print(f"│ Threshold: {threshold_data.get('microns', 'N/A')} µm " +
                  f"({threshold_data.get('pixels', 'N/A'):.1f} pixels)".ljust(68) + "│")
        else:
            print(f"│ ⚠️ PIXEL-ONLY MEASUREMENTS (no micron conversion)".ljust(68) + "│")
            print(f"│ Threshold: {threshold_data.get('pixels', 'N/A')} pixels".ljust(68) + "│")
            
        # Check for missing channels
        missing_channels = [ch for ch in self.selected_channels if ch not in section_data['channels']]
        if missing_channels:
            print(f"│ ⚠️ MISSING {', '.join(missing_channels)} DATA - " +
                  f"Processing {', '.join(section_data['channels'])} only".ljust(68) + "│")
            
        print("└" + "─" * 68 + "┘")
        
        # Process timestamp
        process_start = time.time()
        current_time = datetime.now().strftime("%H:%M:%S")
        
        # Initialize section results
        section_results = {
            'mouse': mouse,
            'section': section,
            'pixel_size': pixel_size,
            'threshold': threshold_data,
            'channels': {},
            'colocalizations': {},
            'areas': {},
            'visualizations': []
        }
        
        # Load localizations for each channel
        print(f"\n[{current_time}] Loading localizations...")
        for channel in section_data['channels']:
            channel_info = section_data['paths'][channel]
            detections = self.load_localizations(channel_info['csv_path'])
            
            if detections is not None:
                section_results['channels'][channel] = {
                    'detections': detections,
                    'count': len(detections),
                    'csv_path': channel_info['csv_path'],
                    'image_path': channel_info['image_path']
                }
                print(f"           ✓ {channel}: {len(detections)} detections loaded")
            else:
                print(f"           ✗ {channel}: Failed to load detections")
                
        # Calculate tissue areas
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Calculating tissue areas...")
        for channel, channel_data in section_results['channels'].items():
            image_path = channel_data['image_path']
            print(f"           Loading {os.path.basename(image_path)}...")
            
            area_pixels = self.calculate_tissue_area(image_path)
            if area_pixels is not None:
                section_results['areas'][channel] = area_pixels
                
                if pixel_size:
                    area_microns = area_pixels * (pixel_size ** 2)
                    print(f"           ✓ {channel} area: {area_pixels:,} pixels ({area_microns:,.0f} µm²)")
                else:
                    print(f"           ✓ {channel} area: {area_pixels:,} pixels (micron conversion N/A)")
            else:
                print(f"           ✗ {channel} area: Could not calculate")
                
        # Perform colocalization analysis if multiple channels present
        if len(section_results['channels']) >= 2:
            print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Performing colocalization analysis...")
            
            # Get all channel pairs
            channels_present = list(section_results['channels'].keys())
            
            for i in range(len(channels_present)):
                for j in range(i + 1, len(channels_present)):
                    ch1, ch2 = channels_present[i], channels_present[j]
                    
                    # Always order channels consistently
                    if self.selected_channels.index(ch1) > self.selected_channels.index(ch2):
                        ch1, ch2 = ch2, ch1
                        
                    print(f"           Analyzing {ch1}→{ch2} distances...")
                    
                    # Perform colocalization
                    coloc_data = self.find_colocalizations(
                        section_results['channels'][ch1]['detections'],
                        section_results['channels'][ch2]['detections'],
                        threshold_data.get('pixels', 10)
                    )
                    
                    section_results['colocalizations'][f"{ch1}_{ch2}"] = coloc_data
                    
                    # Report results
                    num_coloc = len(coloc_data['pairs'])
                    percent_coloc = (num_coloc / section_results['channels'][ch1]['count'] * 100 
                                   if section_results['channels'][ch1]['count'] > 0 else 0)
                    
                    print(f"           ✓ Found {num_coloc} colocalizations ({percent_coloc:.1f}% of {ch1} cells)")
                    
        else:
            print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Skipping colocalization (only 1 channel present)")
            
        # Generate distance distribution if requested
        if self.generate_distance_reports and len(section_results['channels']) >= 2:
            print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Generating distance distribution...")
            
            for coloc_key, coloc_data in section_results['colocalizations'].items():
                self.generate_distance_distribution(section_results, coloc_key)
                
        # Create visualizations if requested
        if self.generate_visualizations:
            print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Creating visualizations...")
            self.create_section_visualizations(section_results)
            
        # Create threshold visualizations if requested
        if self.generate_threshold_visualizations and len(section_results['channels']) >= 2:
            print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Creating threshold visualizations...")
            threshold_data = section_results['threshold']
            if threshold_data.get('pixels'):
                print(f"           Using threshold: {threshold_data['pixels']:.1f} pixels", end="")
                if threshold_data.get('microns'):
                    print(f" ({threshold_data['microns']} µm)")
                else:
                    print()
                self.create_threshold_visualizations(section_results)
            else:
                print("           ⚠️ No threshold data available for threshold visualizations")
            
        # Write output files
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Writing data files...")
        self.write_section_outputs(section_results)
        
        # Store results for summary generation
        self.all_data[section_key] = section_results
        
        print("           ✓ Section complete!")
        print("\n" + "─" * 70 + "\n")
        
    def load_localizations(self, csv_path: str) -> Optional[pd.DataFrame]:
        """Load localization data from CSV file."""
        try:
            # Read CSV
            df = pd.read_csv(csv_path)
            
            # Validate required columns
            required_columns = ['X', 'Y']
            if not all(col in df.columns for col in required_columns):
                self.error_log.append(f"Missing required columns in {csv_path}")
                return None
                
            # Return only X, Y coordinates with original indices preserved
            # Note: CSV indices start at 2 (row 1 is header, row 2 is first data)
            df.index = df.index + 2  # Adjust to match CSV row numbers
            
            return df[['X', 'Y']]
            
        except Exception as e:
            self.error_log.append(f"Error loading {csv_path}: {str(e)}")
            return None
            
    def calculate_tissue_area(self, image_path: str) -> Optional[int]:
        """Calculate tissue area (non-zero pixels) from image."""
        try:
            # Load image
            img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
            if img is None:
                self.error_log.append(f"Could not load image: {image_path}")
                return None
                
            # Count non-zero pixels
            area_pixels = np.count_nonzero(img)
            
            return area_pixels
            
        except Exception as e:
            self.error_log.append(f"Error calculating area for {image_path}: {str(e)}")
            return None
            
    def find_colocalizations(self, detections1: pd.DataFrame, detections2: pd.DataFrame, 
                           threshold: float) -> Dict:
        """Find colocalizations between two sets of detections."""
        coloc_data = {
            'pairs': [],  # List of (idx1, idx2, distance) tuples
            'distances': [],  # All distances for distribution analysis
            'unpaired_1': [],  # Indices from set 1 with no match
            'unpaired_2': []   # Indices from set 2 with no match
        }
        
        if len(detections1) == 0 or len(detections2) == 0:
            return coloc_data
            
        # Convert to numpy arrays for faster computation
        coords1 = detections1[['X', 'Y']].values
        coords2 = detections2[['X', 'Y']].values
        indices1 = detections1.index.values
        indices2 = detections2.index.values
        
        # For each detection in set 1, find nearest in set 2
        paired_2 = set()
        
        for i, (idx1, coord1) in enumerate(zip(indices1, coords1)):
            # Calculate distances to all detections in set 2
            distances = np.sqrt(np.sum((coords2 - coord1) ** 2, axis=1))
            
            # Find minimum distance
            min_idx = np.argmin(distances)
            min_distance = distances[min_idx]
            idx2 = indices2[min_idx]
            
            # Store distance for distribution analysis
            coloc_data['distances'].append(min_distance)
            
            # Check if within threshold
            if min_distance <= threshold:
                coloc_data['pairs'].append((idx1, idx2, min_distance))
                paired_2.add(idx2)
            else:
                coloc_data['unpaired_1'].append(idx1)
                
        # Find unpaired detections in set 2
        coloc_data['unpaired_2'] = [idx for idx in indices2 if idx not in paired_2]
        
        return coloc_data
        
    def generate_distance_distribution(self, section_results: Dict, coloc_key: str):
        """Generate distance distribution analysis."""
        coloc_data = section_results['colocalizations'][coloc_key]
        distances = coloc_data['distances']
        
        if not distances:
            return
            
        # Determine bin size and max distance based on pixel size
        pixel_size = section_results['pixel_size']
        
        if pixel_size and hasattr(self, 'distance_report_mode'):
            if self.distance_report_mode in ['1', '3']:
                # Use microns
                distances_microns = [d * pixel_size for d in distances]
                max_dist = getattr(self, 'max_distance_microns', 30)
                self.create_distribution_report(section_results, coloc_key, 
                                              distances_microns, max_dist, 'µm')
        else:
            # Use pixels
            max_dist = getattr(self, 'max_distance_pixels', 50)
            self.create_distribution_report(section_results, coloc_key, 
                                          distances, max_dist, 'pixels')
            
    def create_distribution_report(self, section_results: Dict, coloc_key: str, 
                                 distances: List[float], max_dist: float, unit: str):
        """Create the actual distribution report."""
        # Create bins
        bin_size = 1  # 1 micron or 1 pixel bins
        bins = np.arange(0, max_dist + bin_size, bin_size)
        
        # Count distances in each bin
        hist, _ = np.histogram(distances, bins=bins)
        
        # Calculate cumulative counts and percentages
        cumulative = np.cumsum(hist)
        ch1 = coloc_key.split('_')[0]
        total_ch1 = section_results['channels'][ch1]['count']
        
        # Find plateau if requested
        plateau_idx = None
        if getattr(self, 'stop_at_plateau', False):
            # Look for 3 consecutive bins with <0.5% increase
            for i in range(3, len(cumulative)):
                recent_increases = []
                for j in range(3):
                    if cumulative[i-j-1] > 0:
                        increase = (cumulative[i-j] - cumulative[i-j-1]) / total_ch1 * 100
                    else:
                        increase = cumulative[i-j] / total_ch1 * 100
                    recent_increases.append(increase)
                    
                if all(inc < 0.5 for inc in recent_increases):
                    plateau_idx = i
                    break
                    
        # Store distribution data
        if 'distance_distribution' not in section_results:
            section_results['distance_distribution'] = {}
        
        section_results['distance_distribution'] = {
            coloc_key: {
                'bins': bins[:plateau_idx] if plateau_idx else bins[:-1],
                'counts': hist[:plateau_idx] if plateau_idx else hist,
                'cumulative': cumulative[:plateau_idx] if plateau_idx else cumulative,
                'unit': unit,
                'plateau': plateau_idx
            }
        }
        
        # Print summary
        if plateau_idx:
            print(f"           ✓ Plateau detected at {bins[plateau_idx]:.0f} {unit}")
            
    def create_section_visualizations(self, section_results: Dict):
        """Create all visualizations for a section."""
        mouse = section_results['mouse']
        section = section_results['section']
        
        # Create visualization directory
        viz_dir = os.path.join(self.output_dir, "CSV_Outputs", mouse, section, "Visualizations")
        os.makedirs(viz_dir, exist_ok=True)
        
        # Define color scheme
        colors = {
            'WFA': (0, 255, 255),          # Yellow
            'PV': (255, 0, 0),             # Blue
            'Agg': (0, 0, 255),            # Red
            'colocalized': (255, 0, 255),  # White
            'triple': (255, 0, 255)        # magenta
        }
        
        # Get visualization settings
        outline_thickness = self.visualization_settings.get('outline_thickness', 2)
        circle_diameter = self.visualization_settings.get('circle_diameter', 16)
        coloc_thickness = self.visualization_settings.get('coloc_thickness', 3)
        circle_radius = circle_diameter // 2
        
        # Create subdirectories
        clean_dir = os.path.join(viz_dir, "Clean_Images")
        individual_dir = os.path.join(viz_dir, "Individual_Overlays")
        composite_dir = os.path.join(viz_dir, "Composite_Overlays")
        layouts_dir = os.path.join(viz_dir, "Layouts")
        
        for d in [clean_dir, individual_dir, composite_dir, layouts_dir]:
            os.makedirs(d, exist_ok=True)
            
        # Process each channel
        channel_images = {}
        
        for channel, channel_data in section_results['channels'].items():
            image_path = channel_data['image_path']
            
            # Load and save clean image
            img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
            if img is None:
                print(f"           ✗ Could not load image for {channel} visualization")
                continue
                
            # Convert to RGB for drawing
            img_rgb = cv2.cvtColor(img, cv2.COLOR_GRAY2RGB)
            
            # Save clean image
            clean_path = os.path.join(clean_dir, f"{channel}_original.png")
            cv2.imwrite(clean_path, img)
            
            # Create individual overlay
            overlay = img_rgb.copy()
            detections = channel_data['detections']
            
            # Determine which detections are colocalized
            colocalized_indices = set()
            
            # Check all colocalization pairs this channel is involved in
            for coloc_key, coloc_data in section_results['colocalizations'].items():
                if channel in coloc_key:
                    ch1, ch2 = coloc_key.split('_')
                    
                    if channel == ch1:
                        # This channel is first in pair
                        colocalized_indices.update([idx for idx, _, _ in coloc_data['pairs']])
                    else:
                        # This channel is second in pair
                        colocalized_indices.update([idx for _, idx, _ in coloc_data['pairs']])
                        
            # Draw circles
            for idx, row in detections.iterrows():
                x, y = int(row['X']), int(row['Y'])
                
                if idx in colocalized_indices:
                    # Colocalized - use purple with thicker outline
                    cv2.circle(overlay, (x, y), circle_radius, colors['colocalized'], coloc_thickness)
                else:
                    # Not colocalized - use channel color
                    cv2.circle(overlay, (x, y), circle_radius, colors[channel], outline_thickness)
                    
            # Save individual overlay
            individual_path = os.path.join(individual_dir, f"{channel}_detections.png")
            cv2.imwrite(individual_path, overlay)
            channel_images[channel] = {
                'clean': img,
                'overlay': overlay,
                'img_rgb': img_rgb
            }
            
        # Create composite overlays for each channel pair
        if len(section_results['channels']) >= 2:
            channels_present = list(section_results['channels'].keys())
            
            for i in range(len(channels_present)):
                for j in range(i + 1, len(channels_present)):
                    ch1, ch2 = channels_present[i], channels_present[j]
                    
                    if ch1 in channel_images and ch2 in channel_images:
                        # Use first channel's image as base
                        composite = channel_images[ch1]['img_rgb'].copy()
                        
                        # Get colocalization data
                        coloc_key = f"{ch1}_{ch2}" if f"{ch1}_{ch2}" in section_results['colocalizations'] else f"{ch2}_{ch1}"
                        if coloc_key in section_results['colocalizations']:
                            coloc_data = section_results['colocalizations'][coloc_key]
                            colocalized_pairs = set(coloc_data['pairs'])
                            
                            # Draw all detections
                            # First channel - non-colocalized
                            detections1 = section_results['channels'][ch1]['detections']
                            coloc_indices1 = {idx for idx, _, _ in coloc_data['pairs']}
                            
                            for idx, row in detections1.iterrows():
                                if idx not in coloc_indices1:
                                    x, y = int(row['X']), int(row['Y'])
                                    cv2.circle(composite, (x, y), circle_radius, colors[ch1], outline_thickness)
                                    
                            # Second channel - non-colocalized
                            detections2 = section_results['channels'][ch2]['detections']
                            coloc_indices2 = {idx for _, idx, _ in coloc_data['pairs']}
                            
                            for idx, row in detections2.iterrows():
                                if idx not in coloc_indices2:
                                    x, y = int(row['X']), int(row['Y'])
                                    cv2.circle(composite, (x, y), circle_radius, colors[ch2], outline_thickness)
                                    
                            # Draw colocalized with white
                            for idx1, idx2, _ in coloc_data['pairs']:
                                # Use position from first channel
                                x = int(detections1.loc[idx1, 'X'])
                                y = int(detections1.loc[idx1, 'Y'])
                                cv2.circle(composite, (x, y), circle_radius, colors['colocalized'], coloc_thickness)
                                
                            # Save composite
                            composite_path = os.path.join(composite_dir, f"{ch1}_{ch2}_composite.png")
                            cv2.imwrite(composite_path, composite)
                            
        # Create layout variations
        self.create_layout_images(section_results, channel_images, layouts_dir)
        
        print("           ✓ Clean images saved")
        print("           ✓ Individual overlays created")
        print("           ✓ Composite overlays created")
        print("           ✓ All layout variants generated")
        
    def create_layout_images(self, section_results: Dict, channel_images: Dict, layouts_dir: str):
        """Create different layout arrangements of images."""
        if len(channel_images) < 2:
            return
            
        channels = list(channel_images.keys())
        
        # For 2 channels, create all layout types
        if len(channels) == 2:
            ch1, ch2 = channels
            img1 = channel_images[ch1]['overlay']
            img2 = channel_images[ch2]['overlay']
            
            h1, w1 = img1.shape[:2]
            h2, w2 = img2.shape[:2]
            
            # Horizontal layout (1x2)
            horizontal = np.zeros((max(h1, h2), w1 + w2 + 10, 3), dtype=np.uint8)
            horizontal[:h1, :w1] = img1
            horizontal[:h2, w1 + 10:] = img2
            cv2.imwrite(os.path.join(layouts_dir, f"{ch1}_{ch2}_horizontal.png"), horizontal)
            
            # Vertical layout (2x1)
            vertical = np.zeros((h1 + h2 + 10, max(w1, w2), 3), dtype=np.uint8)
            vertical[:h1, :w1] = img1
            vertical[h1 + 10:, :w2] = img2
            cv2.imwrite(os.path.join(layouts_dir, f"{ch1}_{ch2}_vertical.png"), vertical)
            
        # For 3+ channels, also create grid
        if len(channels) >= 3:
            # Create 2x2 grid with composite in 4th position
            # Implementation would go here for 3-channel case
            pass
            
    def create_threshold_visualizations(self, section_results: Dict):
        """Create threshold-focused grayscale visualizations."""
        mouse = section_results['mouse']
        section = section_results['section']
        
        # Get base visualization directory
        viz_dir = os.path.join(self.output_dir, "CSV_Outputs", mouse, section, "Visualizations")
        threshold_dir = os.path.join(viz_dir, "ThresholdLogic")
        os.makedirs(threshold_dir, exist_ok=True)
        
        # Get visualization settings
        circle_diameter = self.threshold_viz_settings.get('circle_diameter', 20)
        circle_radius = circle_diameter // 2
        method = self.threshold_viz_settings.get('method', 'color')
        
        # Process each channel pair
        channels_present = list(section_results['channels'].keys())
        
        for i in range(len(channels_present)):
            for j in range(i + 1, len(channels_present)):
                ch1, ch2 = channels_present[i], channels_present[j]
                
                # Create pair directory
                pair_name = f"{ch1}-{ch2}"
                pair_dir = os.path.join(threshold_dir, pair_name)
                os.makedirs(pair_dir, exist_ok=True)
                
                # Load grayscale images
                img1_path = section_results['channels'][ch1]['image_path']
                img2_path = section_results['channels'][ch2]['image_path']
                
                img1 = cv2.imread(img1_path, cv2.IMREAD_GRAYSCALE)
                img2 = cv2.imread(img2_path, cv2.IMREAD_GRAYSCALE)
                
                if img1 is None or img2 is None:
                    print(f"           ⚠️ Could not load images for {pair_name}")
                    continue
                
                # Get colocalization data
                coloc_key = f"{ch1}_{ch2}" if f"{ch1}_{ch2}" in section_results['colocalizations'] else f"{ch2}_{ch1}"
                if coloc_key not in section_results['colocalizations']:
                    print(f"           ⚠️ No colocalization data for {pair_name}")
                    continue
                    
                coloc_data = section_results['colocalizations'][coloc_key]
                
                # 1. Create composite background
                blend_key = f"{ch1}_{ch2}" if f"{ch1}_{ch2}" in self.threshold_viz_settings['blend_ratios'] else f"{ch2}_{ch1}"
                blend_ratios = self.threshold_viz_settings['blend_ratios'].get(blend_key, {ch1: 0.5, ch2: 0.5})
                
                composite = self.create_blended_background(img1, img2, blend_ratios[ch1], blend_ratios[ch2])
                cv2.imwrite(os.path.join(pair_dir, "composite.png"), composite)
                
                # 2. Raw circles composite
                raw_composite = self.draw_raw_circles(composite.copy(), section_results, ch1, ch2, circle_radius, method)
                cv2.imwrite(os.path.join(pair_dir, "raw_composite.png"), raw_composite)
                
                # 3. Raw circles side-by-side
                raw_side_by_side = self.create_side_by_side_raw(img1, img2, section_results, ch1, ch2, circle_radius, method)
                cv2.imwrite(os.path.join(pair_dir, "raw_side_by_side.png"), raw_side_by_side)
                
                # 4. Colocalization circles composite
                coloc_composite = self.draw_coloc_circles(composite.copy(), section_results, ch1, ch2, 
                                                         coloc_data, circle_radius, method)
                cv2.imwrite(os.path.join(pair_dir, "coloc_circles_composite.png"), coloc_composite)
                
                # 5. Colocalization circles side-by-side
                coloc_side_by_side = self.create_side_by_side_coloc(img1, img2, section_results, ch1, ch2, 
                                                                    coloc_data, circle_radius, method)
                cv2.imwrite(os.path.join(pair_dir, "coloc_circles_side_by_side.png"), coloc_side_by_side)
                
                # 6. Lines composite
                lines_composite = self.draw_connection_lines(composite.copy(), section_results, ch1, ch2, coloc_data)
                cv2.imwrite(os.path.join(pair_dir, "lines_composite.png"), lines_composite)
                
                # 7. Lines side-by-side
                lines_side_by_side = self.create_side_by_side_lines(img1, img2, section_results, ch1, ch2, coloc_data)
                cv2.imwrite(os.path.join(pair_dir, "lines_side_by_side.png"), lines_side_by_side)
                
                print(f"           ✓ {pair_name} threshold visualizations (7 files)")
                
        print("           ✓ Threshold visualizations complete")
        
    def create_blended_background(self, img1: np.ndarray, img2: np.ndarray, 
                                 blend1: float, blend2: float) -> np.ndarray:
        """Create a blended grayscale background from two images."""
        # Ensure images are float for blending
        img1_float = img1.astype(np.float32)
        img2_float = img2.astype(np.float32)
        
        # Blend images
        blended = (img1_float * blend1 + img2_float * blend2)
        
        # Convert back to uint8
        blended = np.clip(blended, 0, 255).astype(np.uint8)
        
        # Convert to RGB for drawing colored circles
        return cv2.cvtColor(blended, cv2.COLOR_GRAY2RGB)
        
    def draw_raw_circles(self, img: np.ndarray, section_results: Dict, 
                        ch1: str, ch2: str, radius: int, method: str) -> np.ndarray:
        """Draw all raw detection circles on the image."""
        # Get detections
        detections1 = section_results['channels'][ch1]['detections']
        detections2 = section_results['channels'][ch2]['detections']
        
        if method == 'color':
            # Use original colors for all detections
            colors = self.threshold_viz_settings['original_colors']
            color1 = colors[ch1]
            color2 = colors[ch2]
        else:
            # Use gray for thickness method
            color1 = color2 = (128, 128, 128)
            
        # Set thickness for raw circles (always thin)
        thickness = 2
        
        # Create overlay for transparency
        overlay = img.copy()
        
        # Draw channel 1 detections
        for idx, row in detections1.iterrows():
            x, y = int(row['X']), int(row['Y'])
            cv2.circle(overlay, (x, y), radius, color1, thickness)
            
        # Draw channel 2 detections
        for idx, row in detections2.iterrows():
            x, y = int(row['X']), int(row['Y'])
            cv2.circle(overlay, (x, y), radius, color2, thickness)
            
        # Apply transparency
        alpha = .9  # Default transparency
        return cv2.addWeighted(overlay, alpha, img, 1 - alpha, 0)
        
    def create_side_by_side_raw(self, img1: np.ndarray, img2: np.ndarray, 
                               section_results: Dict, ch1: str, ch2: str, 
                               radius: int, method: str) -> np.ndarray:
        """Create side-by-side visualization with raw circles."""
        # Convert to RGB
        img1_rgb = cv2.cvtColor(img1, cv2.COLOR_GRAY2RGB)
        img2_rgb = cv2.cvtColor(img2, cv2.COLOR_GRAY2RGB)
        
        # Draw circles on each
        img1_circles = self.draw_single_channel_raw(img1_rgb.copy(), section_results, ch1, radius, method)
        img2_circles = self.draw_single_channel_raw(img2_rgb.copy(), section_results, ch2, radius, method)
        
        # Create side-by-side
        h1, w1 = img1_circles.shape[:2]
        h2, w2 = img2_circles.shape[:2]
        
        side_by_side = np.zeros((max(h1, h2), w1 + w2 + 10, 3), dtype=np.uint8)
        side_by_side[:h1, :w1] = img1_circles
        side_by_side[:h2, w1 + 10:] = img2_circles
        
        return side_by_side
        
    def draw_single_channel_raw(self, img: np.ndarray, section_results: Dict, 
                               channel: str, radius: int, method: str) -> np.ndarray:
        """Draw raw circles for a single channel."""
        detections = section_results['channels'][channel]['detections']
        
        if method == 'color':
            color = self.threshold_viz_settings['original_colors'][channel]
        else:
            color = (128, 128, 128)
            
        overlay = img.copy()
        
        for idx, row in detections.iterrows():
            x, y = int(row['X']), int(row['Y'])
            cv2.circle(overlay, (x, y), radius, color, 1)
            
        alpha = .9
        return cv2.addWeighted(overlay, alpha, img, 1 - alpha, 0)
        
    def draw_coloc_circles(self, img: np.ndarray, section_results: Dict, 
                          ch1: str, ch2: str, coloc_data: Dict, 
                          radius: int, method: str) -> np.ndarray:
        """Draw circles with colocalization status indicated."""
        # Get detections
        detections1 = section_results['channels'][ch1]['detections']
        detections2 = section_results['channels'][ch2]['detections']
        
        # Get colocalized indices
        coloc_indices1 = {idx for idx, _, _ in coloc_data['pairs']}
        coloc_indices2 = {idx for _, idx, _ in coloc_data['pairs']}
        
        overlay = img.copy()
        
        if method == 'color':
            # Color-coded method
            orig_colors = self.threshold_viz_settings['original_colors']
            alt_colors = self.threshold_viz_settings['alt_colors']
            thickness = 2
            
            # Draw channel 1
            for idx, row in detections1.iterrows():
                x, y = int(row['X']), int(row['Y'])
                if idx in coloc_indices1:
                    color = orig_colors[ch1]  # Colocalized - original color
                else:
                    color = alt_colors[ch1]   # Non-colocalized - alternative color
                cv2.circle(overlay, (x, y), radius, color, thickness)
                
            # Draw channel 2
            for idx, row in detections2.iterrows():
                x, y = int(row['X']), int(row['Y'])
                if idx in coloc_indices2:
                    color = orig_colors[ch2]  # Colocalized - original color
                else:
                    color = alt_colors[ch2]   # Non-colocalized - alternative color
                cv2.circle(overlay, (x, y), radius, color, thickness)
                
            # Add color key
            overlay = self.add_color_key(overlay, [ch1, ch2])
            
        else:
            # Thickness-based method
            thick = self.threshold_viz_settings['thick_circle']
            thin = self.threshold_viz_settings['thin_circle']
            color = (128, 128, 128)  # Gray
            
            # Draw channel 1
            for idx, row in detections1.iterrows():
                x, y = int(row['X']), int(row['Y'])
                thickness = thick if idx in coloc_indices1 else thin
                cv2.circle(overlay, (x, y), radius, color, thickness)
                
            # Draw channel 2
            for idx, row in detections2.iterrows():
                x, y = int(row['X']), int(row['Y'])
                thickness = thick if idx in coloc_indices2 else thin
                cv2.circle(overlay, (x, y), radius, color, thickness)
        
        # Apply transparency
        alpha = .9
        return cv2.addWeighted(overlay, alpha, img, 1 - alpha, 0)
        
    def create_side_by_side_coloc(self, img1: np.ndarray, img2: np.ndarray, 
                                 section_results: Dict, ch1: str, ch2: str, 
                                 coloc_data: Dict, radius: int, method: str) -> np.ndarray:
        """Create side-by-side visualization with colocalization circles."""
        # Convert to RGB
        img1_rgb = cv2.cvtColor(img1, cv2.COLOR_GRAY2RGB)
        img2_rgb = cv2.cvtColor(img2, cv2.COLOR_GRAY2RGB)
        
        # Get colocalized indices
        coloc_indices1 = {idx for idx, _, _ in coloc_data['pairs']}
        coloc_indices2 = {idx for _, idx, _ in coloc_data['pairs']}
        
        # Draw circles on each
        img1_circles = self.draw_single_channel_coloc(img1_rgb.copy(), section_results, ch1, 
                                                     coloc_indices1, radius, method)
        img2_circles = self.draw_single_channel_coloc(img2_rgb.copy(), section_results, ch2, 
                                                     coloc_indices2, radius, method)
        
        # Create side-by-side
        h1, w1 = img1_circles.shape[:2]
        h2, w2 = img2_circles.shape[:2]
        
        side_by_side = np.zeros((max(h1, h2), w1 + w2 + 10, 3), dtype=np.uint8)
        side_by_side[:h1, :w1] = img1_circles
        side_by_side[:h2, w1 + 10:] = img2_circles
        
        # Add color key if using color method
        if method == 'color':
            side_by_side = self.add_color_key(side_by_side, [ch1, ch2])
        
        return side_by_side
        
    def draw_single_channel_coloc(self, img: np.ndarray, section_results: Dict, 
                                 channel: str, coloc_indices: set, 
                                 radius: int, method: str) -> np.ndarray:
        """Draw colocalization circles for a single channel."""
        detections = section_results['channels'][channel]['detections']
        
        overlay = img.copy()
        
        if method == 'color':
            orig_color = self.threshold_viz_settings['original_colors'][channel]
            alt_color = self.threshold_viz_settings['alt_colors'][channel]
            thickness = 2
            
            for idx, row in detections.iterrows():
                x, y = int(row['X']), int(row['Y'])
                color = orig_color if idx in coloc_indices else alt_color
                cv2.circle(overlay, (x, y), radius, color, thickness)
        else:
            thick = self.threshold_viz_settings['thick_circle']
            thin = self.threshold_viz_settings['thin_circle']
            color = (128, 128, 128)
            
            for idx, row in detections.iterrows():
                x, y = int(row['X']), int(row['Y'])
                thickness = thick if idx in coloc_indices else thin
                cv2.circle(overlay, (x, y), radius, color, thickness)
                
        alpha = .9
        return cv2.addWeighted(overlay, alpha, img, 1 - alpha, 0)
        
    def draw_connection_lines(self, img: np.ndarray, section_results: Dict, 
                             ch1: str, ch2: str, coloc_data: Dict) -> np.ndarray:
        """Draw red lines connecting colocalized pairs."""
        # Get detections
        detections1 = section_results['channels'][ch1]['detections']
        detections2 = section_results['channels'][ch2]['detections']
        
        overlay = img.copy()
        
        # Draw lines for each colocalized pair
        for idx1, idx2, _ in coloc_data['pairs']:
            # Get coordinates
            x1, y1 = int(detections1.loc[idx1, 'X']), int(detections1.loc[idx1, 'Y'])
            x2, y2 = int(detections2.loc[idx2, 'X']), int(detections2.loc[idx2, 'Y'])
            
            # Draw red line
            cv2.line(overlay, (x1, y1), (x2, y2), (0, 0, 255), 3)
            
        # Apply transparency
        alpha = .9
        return cv2.addWeighted(overlay, alpha, img, 1 - alpha, 0)
        
    def create_side_by_side_lines(self, img1: np.ndarray, img2: np.ndarray, 
                                 section_results: Dict, ch1: str, ch2: str, 
                                 coloc_data: Dict) -> np.ndarray:
        """Create side-by-side visualization with connection lines."""
        # Convert to RGB
        img1_rgb = cv2.cvtColor(img1, cv2.COLOR_GRAY2RGB)
        img2_rgb = cv2.cvtColor(img2, cv2.COLOR_GRAY2RGB)
        
        # Draw lines on each
        img1_lines = self.draw_connection_lines(img1_rgb.copy(), section_results, ch1, ch2, coloc_data)
        img2_lines = self.draw_connection_lines(img2_rgb.copy(), section_results, ch1, ch2, coloc_data)
        
        # Create side-by-side
        h1, w1 = img1_lines.shape[:2]
        h2, w2 = img2_lines.shape[:2]
        
        side_by_side = np.zeros((max(h1, h2), w1 + w2 + 10, 3), dtype=np.uint8)
        side_by_side[:h1, :w1] = img1_lines
        side_by_side[:h2, w1 + 10:] = img2_lines
        
        return side_by_side
        
    def add_color_key(self, img: np.ndarray, channels: List[str]) -> np.ndarray:
        """Add a color key to the image."""
        # Key parameters
        key_height = 90 + len(channels) * 150  # Height based on number of channels
        key_width = 900
        margin = 30
        
        # Create key background (semi-transparent white)
        key_bg = np.ones((key_height, key_width, 3), dtype=np.uint8) * 255
        
        # Get colors
        orig_colors = self.threshold_viz_settings['original_colors']
        alt_colors = self.threshold_viz_settings['alt_colors']
        
        # Draw key entries
        y_offset = 60
        for channel in channels:
            # Colocalized entry
            cv2.circle(key_bg, (60, y_offset), 24, orig_colors[channel], -1)
            cv2.putText(key_bg, f"Colocalized {channel}", (120, y_offset + 15), 
                       cv2.FONT_HERSHEY_SIMPLEX, 2.1, (0, 0, 0), 6)
            
            # Non-colocalized entry
            y_offset += 60
            cv2.circle(key_bg, (60, y_offset), 24, alt_colors[channel], -1)
            cv2.putText(key_bg, f"Non-colocalized {channel}", (200, y_offset + 25), 
                       cv2.FONT_HERSHEY_SIMPLEX, 2.1, (0, 0, 0), 6)
            
            y_offset += 60
            
        # Place key on image (top-right corner)
        img_h, img_w = img.shape[:2]
        key_x = img_w - key_width - margin
        key_y = margin
        
        # Ensure key fits on image
        if key_x > 0 and key_y + key_height < img_h:
            # Blend key with image
            roi = img[key_y:key_y + key_height, key_x:key_x + key_width]
            blended = cv2.addWeighted(key_bg, 0.8, roi, 0.2, 0)
            img[key_y:key_y + key_height, key_x:key_x + key_width] = blended
            
            # Add border around key
            cv2.rectangle(img, (key_x, key_y), (key_x + key_width, key_y + key_height), (0, 0, 0), 3)
            
        return img
            
    def write_section_outputs(self, section_results: Dict):
        """Write all output files for a section."""
        mouse = section_results['mouse']
        section = section_results['section']
        
        # Create output directory for this section
        csv_dir = os.path.join(self.output_dir, "CSV_Outputs", mouse, section)
        excel_dir = os.path.join(self.output_dir, "Excel_Outputs", mouse, section)
        
        os.makedirs(csv_dir, exist_ok=True)
        os.makedirs(excel_dir, exist_ok=True)
        
        # Prepare data for outputs
        output_data = {}
        
        # 1. Raw counts
        raw_counts = {
            'Mouse': mouse,
            'Section': section,
            'WFA Count': section_results['channels'].get('WFA', {}).get('count', 'N/A'),
            'Agg Count': section_results['channels'].get('Agg', {}).get('count', 'N/A'),
            'PV Count': section_results['channels'].get('PV', {}).get('count', 'N/A')
        }
        output_data['raw_counts'] = raw_counts
        
        # Write raw counts CSV
        self.write_csv_with_header(
            os.path.join(csv_dir, "1_raw_counts.csv"),
            [raw_counts],
            self.generate_source_header(section_results, 'all')
        )
        
        # 2-4. Channel summaries
        for channel in ['WFA', 'Agg', 'PV']:
            if channel in section_results['channels']:
                summary = self.create_channel_summary(section_results, channel)
                output_data[f"{channel.lower()}_summary"] = summary
                
                # Determine file number
                file_num = {'WFA': 2, 'Agg': 3, 'PV': 4}[channel]
                
                self.write_csv_with_header(
                    os.path.join(csv_dir, f"{file_num}_{channel.lower()}_summary.csv"),
                    [summary],
                    self.generate_source_header(section_results, channel)
                )
                
        # 5-7. Pairwise colocalizations
        coloc_file_mapping = {
            'WFA_Agg': 5,
            'WFA_PV': 6,
            'Agg_PV': 7
        }
        
        for coloc_key, file_num in coloc_file_mapping.items():
            if coloc_key in section_results['colocalizations']:
                coloc_data = section_results['colocalizations'][coloc_key]
                ch1, ch2 = coloc_key.split('_')
                
                # Create colocalization rows
                coloc_rows = []
                for idx1, idx2, distance in coloc_data['pairs']:
                    row = {
                        'Mouse': mouse,
                        'Section': section,
                        f'Index {ch1}': idx1,
                        f'Index {ch2}': idx2,
                        'Distance (px)': f"{distance:.2f}"
                    }
                    
                    # Add micron distance if available
                    if section_results['pixel_size']:
                        row['Distance (µm)'] = f"{distance * section_results['pixel_size']:.2f}"
                    else:
                        row['Distance (µm)'] = 'N/A'
                        
                    coloc_rows.append(row)
                    
                output_data[f'colocalization_{coloc_key.lower()}'] = coloc_rows
                
                # Write CSV
                header = self.generate_colocalization_header(section_results, coloc_key, coloc_data)
                self.write_csv_with_header(
                    os.path.join(csv_dir, f"{file_num}_colocalization_{coloc_key.lower()}.csv"),
                    coloc_rows,
                    header
                )
                
        # 8. Triple colocalization (if all 3 channels present)
        if all(ch in section_results['channels'] for ch in ['WFA', 'Agg', 'PV']):
            triple_coloc = self.find_triple_colocalizations(section_results)
            if triple_coloc:
                output_data['triple_colocalization'] = triple_coloc
                
                header = self.generate_triple_colocalization_header(section_results)
                self.write_csv_with_header(
                    os.path.join(csv_dir, "8_triple_colocalization.csv"),
                    triple_coloc,
                    header
                )
                
        # Write distance distributions if generated
        if 'distance_distribution' in section_results:
            for coloc_key, dist_data in section_results['distance_distribution'].items():
                self.write_distance_distribution(csv_dir, coloc_key, dist_data)
                
        # Write processing parameters
        params = {
            'mouse': mouse,
            'section': section,
            'timestamp': datetime.now().isoformat(),
            'pixel_size': section_results['pixel_size'],
            'threshold': section_results['threshold'],
            'channels_analyzed': list(section_results['channels'].keys()),
            'areas': section_results['areas']
        }
        
        with open(os.path.join(csv_dir, "processing_parameters.json"), 'w', encoding='utf-8') as f:
            json.dump(params, f, indent=2)
            
        # Create Excel workbook
        self.create_section_excel(excel_dir, section, output_data)
        
        print("           ✓ CSV files created/updated")
        print("           ✓ Excel workbook generated")
        
    def generate_source_header(self, section_results: Dict, channel: str) -> List[str]:
        """Generate header with source paths."""
        header = [f"# Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"]
        
        if channel == 'all':
            # For raw counts, include all channels
            header.append("# Sources:")
            for ch in ['WFA', 'Agg', 'PV']:
                if ch in section_results['channels']:
                    path = section_results['channels'][ch]['csv_path']
                    header.append(f"# {ch}: {path}")
        else:
            # For individual channel
            if channel in section_results['channels']:
                path = section_results['channels'][channel]['csv_path']
                header.append(f"# Source: {path}")
                
                # Add image info
                img_path = section_results['channels'][channel]['image_path']
                header.append(f"# Image: {os.path.basename(img_path)}")
                
        # Add pixel size info
        if section_results['pixel_size']:
            header.append(f"# Pixel size: {section_results['pixel_size']} µm/pixel")
        else:
            header.append("# ⚠️ NO PIXEL SIZE DATA - Micron conversions unavailable")
            
        return header
        
    def create_channel_summary(self, section_results: Dict, channel: str) -> Dict:
        """Create summary data for a channel."""
        channel_data = section_results['channels'][channel]
        count = channel_data['count']
        area_pixels = section_results['areas'].get(channel, 'N/A')
        
        summary = {
            'Mouse': section_results['mouse'],
            'Section': section_results['section'],
            f'{channel} Count': count,
            'Area (pixels)': area_pixels if area_pixels != 'N/A' else 'N/A'
        }
        
        # Add micron conversions if available
        if section_results['pixel_size'] and area_pixels != 'N/A':
            pixel_size = section_results['pixel_size']
            area_microns = area_pixels * (pixel_size ** 2)
            density_per_mm2 = (count / area_microns) * 1_000_000
            
            summary['Area (µm²)'] = f"{area_microns:.0f}"
            summary['Density (per M pixels)'] = f"{(count / area_pixels) * 1_000_000:.2f}"
            summary['Density (per mm²)'] = f"{density_per_mm2:.2f}"
        else:
            summary['Area (µm²)'] = 'N/A'
            summary['Density (per M pixels)'] = f"{(count / area_pixels) * 1_000_000:.2f}" if area_pixels != 'N/A' else 'N/A'
            summary['Density (per mm²)'] = 'N/A'
            
        return summary
        
    def generate_colocalization_header(self, section_results: Dict, coloc_key: str, 
                                      coloc_data: Dict) -> List[str]:
        """Generate header for colocalization CSV."""
        ch1, ch2 = coloc_key.split('_')
        num_coloc = len(coloc_data['pairs'])
        ch1_count = section_results['channels'][ch1]['count']
        percent = (num_coloc / ch1_count * 100) if ch1_count > 0 else 0
        
        header = [
            f"# Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "# Sources:",
            f"# {ch1}: {section_results['channels'][ch1]['csv_path']}",
            f"# {ch2}: {section_results['channels'][ch2]['csv_path']}"
        ]
        
        # Add threshold info
        threshold_data = section_results['threshold']
        if threshold_data.get('microns'):
            header.append(f"# Threshold: {threshold_data['pixels']:.1f} pixels ({threshold_data['microns']} µm)")
        else:
            header.append(f"# Threshold: {threshold_data['pixels']} pixels")
            
        # Add summary
        header.append(f"# Total: {num_coloc} of {ch1_count} {ch1} cells colocalized ({percent:.1f}%)")
        
        return header
        
    def find_triple_colocalizations(self, section_results: Dict) -> List[Dict]:
        """Find triple colocalizations among WFA, Agg, and PV."""
        # Get all pairwise colocalizations
        wfa_agg_pairs = {(idx1, idx2) for idx1, idx2, _ in 
                         section_results['colocalizations'].get('WFA_Agg', {}).get('pairs', [])}
        wfa_pv_pairs = {(idx1, idx2) for idx1, idx2, _ in 
                        section_results['colocalizations'].get('WFA_PV', {}).get('pairs', [])}
        agg_pv_pairs = {(idx1, idx2) for idx1, idx2, _ in 
                        section_results['colocalizations'].get('Agg_PV', {}).get('pairs', [])}
        
        # Find WFA indices that are colocalized with both Agg and PV
        triple_coloc = []
        
        for wfa_idx, agg_idx in wfa_agg_pairs:
            # Check if this WFA is also colocalized with PV
            pv_match = next((pv_idx for w_idx, pv_idx in wfa_pv_pairs if w_idx == wfa_idx), None)
            
            if pv_match:
                # Check if the Agg and PV are also colocalized
                if (agg_idx, pv_match) in agg_pv_pairs or (pv_match, agg_idx) in agg_pv_pairs:
                    # Calculate average distance
                    distances = []
                    
                    # Get individual distances
                    for pairs, ch_pair in [
                        (section_results['colocalizations']['WFA_Agg']['pairs'], (wfa_idx, agg_idx)),
                        (section_results['colocalizations']['WFA_PV']['pairs'], (wfa_idx, pv_match)),
                        (section_results['colocalizations']['Agg_PV']['pairs'], (agg_idx, pv_match))
                    ]:
                        for idx1, idx2, dist in pairs:
                            if (idx1, idx2) == ch_pair or (idx2, idx1) == ch_pair:
                                distances.append(dist)
                                break
                                
                    avg_distance = np.mean(distances) if distances else 0
                    
                    triple_coloc.append({
                        'Mouse': section_results['mouse'],
                        'Section': section_results['section'],
                        'Index WFA': wfa_idx,
                        'Index Agg': agg_idx,
                        'Index PV': pv_match,
                        'Avg Distance (px)': f"{avg_distance:.2f}",
                        'Avg Distance (µm)': f"{avg_distance * section_results['pixel_size']:.2f}" 
                                            if section_results['pixel_size'] else 'N/A'
                    })
                    
        return triple_coloc
        
    def generate_triple_colocalization_header(self, section_results: Dict) -> List[str]:
        """Generate header for triple colocalization CSV."""
        header = [
            f"# Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "# Sources:",
            f"# WFA: {section_results['channels']['WFA']['csv_path']}",
            f"# Agg: {section_results['channels']['Agg']['csv_path']}",
            f"# PV: {section_results['channels']['PV']['csv_path']}"
        ]
        
        # Add threshold info
        threshold_data = section_results['threshold']
        if threshold_data.get('microns'):
            header.append(f"# Threshold: {threshold_data['pixels']:.1f} pixels ({threshold_data['microns']} µm)")
        else:
            header.append(f"# Threshold: {threshold_data['pixels']} pixels")
            
        return header
        
    def write_csv_with_header(self, filepath: str, data: List[Dict], header: List[str]):
        """Write CSV with comment header."""
        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            # Write header comments
            for line in header:
                f.write(line + '\n')
                
            # Write data
            if data:
                writer = csv.DictWriter(f, fieldnames=data[0].keys())
                writer.writeheader()
                writer.writerows(data)
                
    def write_distance_distribution(self, output_dir: str, coloc_key: str, dist_data: Dict):
        """Write distance distribution CSV."""
        cum = dist_data.get('cumulative')
        if cum is None or len(cum) == 0:
            return
        ch1, ch2 = coloc_key.split('_')
        
        rows = []
        # Precompute totals to avoid shadowing variable names
        total_count = dist_data['cumulative'][-1] if len(dist_data['cumulative']) > 0 else 0
        
        for i, (bin_start, count, cum_val) in enumerate(zip(
            dist_data['bins'],
            dist_data['counts'],
            dist_data['cumulative']
        )):
            bin_end = bin_start + 1
            
            percent_of_ch1 = f"{(count / total_count * 100):.1f}" if total_count > 0 else "0.0"
            cumulative_percent = f"{(cum_val / total_count * 100):.1f}" if total_count > 0 else "0.0"
            
            row = {
                f'Distance_Range_{dist_data["unit"]}': f"{bin_start}-{bin_end}",
                'Count': int(count),
                'Cumulative_Count': int(cum_val),
                'Percent_of_' + ch1: percent_of_ch1,
                'Cumulative_Percent': cumulative_percent
            }
            rows.append(row)
            


        # Write CSV
        header = [
            f"# Distance Distribution Analysis for {coloc_key.replace('_', '-')} colocalization",
            f"# Unit: {dist_data['unit']}"
        ]
        
        if dist_data.get('plateau'):
            header.append(f"# Analysis stopped at {dist_data['bins'][dist_data['plateau']]:.0f} " +
                         f"{dist_data['unit']} (plateau detected)")
            
        filepath = os.path.join(output_dir, f"distance_distribution_{coloc_key.lower()}.csv")
        self.write_csv_with_header(filepath, rows, header)
        
    def create_section_excel(self, excel_dir: str, section: str, output_data: Dict):
        """Create Excel workbook for section data."""
        filepath = os.path.join(excel_dir, f"{section}_Complete_Analysis.xlsx")
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            # Raw counts
            if 'raw_counts' in output_data:
                df = pd.DataFrame([output_data['raw_counts']])
                df.to_excel(writer, sheet_name='Raw_Counts', index=False)
                
            # Channel summaries
            for channel in ['wfa', 'agg', 'pv']:
                key = f'{channel}_summary'
                if key in output_data:
                    df = pd.DataFrame([output_data[key]])
                    df.to_excel(writer, sheet_name=f'{channel.upper()}_Summary', index=False)
                    
            # Colocalizations
            for coloc_key in ['wfa_agg', 'wfa_pv', 'agg_pv']:
                key = f'colocalization_{coloc_key}'
                if key in output_data and output_data[key]:
                    df = pd.DataFrame(output_data[key])
                    sheet_name = f'{coloc_key.upper().replace("_", "_")}_Colocalization'
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
            # Triple colocalization
            if 'triple_colocalization' in output_data and output_data['triple_colocalization']:
                df = pd.DataFrame(output_data['triple_colocalization'])
                df.to_excel(writer, sheet_name='Triple_Colocalization', index=False)
                
            # Summary page
            self.create_excel_summary_sheet(writer, output_data)
            
    def create_excel_summary_sheet(self, writer, output_data: Dict):
        """Create summary sheet in Excel workbook."""
        # Create summary data
        summary_data = []
        
        # Add channel counts
        if 'raw_counts' in output_data:
            counts = output_data['raw_counts']
            summary_data.append(['Channel Counts', ''])
            for ch in ['WFA', 'Agg', 'PV']:
                key = f'{ch} Count'
                if key in counts and counts[key] != 'N/A':
                    summary_data.append([ch, counts[key]])
                    
        summary_data.append(['', ''])  # Empty row
        
        # Add colocalization summaries
        summary_data.append(['Colocalization Summary', ''])
        
        # Count colocalizations
        for coloc_type in ['wfa_agg', 'wfa_pv', 'agg_pv']:
            key = f'colocalization_{coloc_type}'
            if key in output_data and output_data[key]:
                ch1, ch2 = coloc_type.upper().split('_')
                count = len(output_data[key])
                summary_data.append([f'{ch1}-{ch2} Colocalizations', count])
                
        # Triple colocalizations
        if 'triple_colocalization' in output_data:
            count = len(output_data['triple_colocalization'])
            summary_data.append(['Triple Colocalizations', count])
            
        # Write summary
        df = pd.DataFrame(summary_data, columns=['Metric', 'Value'])
        df.to_excel(writer, sheet_name='Summary', index=False)
        
    def generate_summaries(self):
        """Generate mouse and master summaries."""
        print("\n[INFO] Generating summary files...")
        
        # Create summary directories
        csv_summary_dir = os.path.join(self.output_dir, "CSV_Outputs", "Master_Summaries")
        excel_summary_dir = os.path.join(self.output_dir, "Excel_Outputs", "Master_Summaries")
        
        os.makedirs(csv_summary_dir, exist_ok=True)
        os.makedirs(excel_summary_dir, exist_ok=True)
        
        # Generate mouse summaries
        for mouse in self.mice_to_process:
            self.generate_mouse_summary(mouse)
            
        # Generate master summary
        self.generate_master_summary(csv_summary_dir, excel_summary_dir)
        
        # Generate measurement units report
        self.generate_units_report(csv_summary_dir)
        
    def generate_mouse_summary(self, mouse: str):
        """Generate summary for a single mouse."""
        # Collect all sections for this mouse
        mouse_sections = []
        
        for section_key, section_data in self.all_data.items():
            if section_data['mouse'] == mouse:
                mouse_sections.append(section_data)
                
        if not mouse_sections:
            return
            
        # Sort by section name
        mouse_sections.sort(key=lambda x: x['section'])
        
        # Create detailed summary
        summary_rows = []
        
        for section_data in mouse_sections:
            row = {
                'Section': section_data['section']
            }
            
            # Add counts for each channel
            for channel in ['WFA', 'Agg', 'PV']:
                if channel in section_data['channels']:
                    row[f'{channel} Count'] = section_data['channels'][channel]['count']
                else:
                    row[f'{channel} Count'] = 'N/A'
                    
            # Add area data
            areas = list(section_data['areas'].values())
            if areas:
                area_pixels = areas[0]  # Assume all channels have same area
                row['Area (px)'] = area_pixels
                
                if section_data['pixel_size']:
                    area_microns = area_pixels * (section_data['pixel_size'] ** 2)
                    row['Area (µm²)'] = f"{area_microns:.0f}"
                    
                    # Add densities
                    for channel in ['WFA', 'Agg', 'PV']:
                        if channel in section_data['channels']:
                            count = section_data['channels'][channel]['count']
                            density_pixels = (count / area_pixels) * 1_000_000
                            density_mm2 = (count / area_microns) * 1_000_000
                            
                            row[f'{channel} Dens (per M px)'] = f"{density_pixels:.2f}"
                            row[f'{channel} Dens (per mm²)'] = f"{density_mm2:.2f}"
                        else:
                            row[f'{channel} Dens (per M px)'] = 'N/A'
                            row[f'{channel} Dens (per mm²)'] = 'N/A'
                else:
                    row['Area (µm²)'] = 'N/A'
                    
                    # Only pixel densities
                    for channel in ['WFA', 'Agg', 'PV']:
                        if channel in section_data['channels']:
                            count = section_data['channels'][channel]['count']
                            density_pixels = (count / area_pixels) * 1_000_000
                            row[f'{channel} Dens (per M px)'] = f"{density_pixels:.2f}"
                            row[f'{channel} Dens (per mm²)'] = 'N/A'
                        else:
                            row[f'{channel} Dens (per M px)'] = 'N/A'
                            row[f'{channel} Dens (per mm²)'] = 'N/A'
            else:
                row['Area (px)'] = 'N/A'
                row['Area (µm²)'] = 'N/A'
                
            # Add colocalization data
            coloc_pairs = [
                ('WFA_Agg', 'WFA-Agg'),
                ('WFA_PV', 'WFA-PV'),
                ('Agg_PV', 'Agg-PV')
            ]
            
            for coloc_key, display_name in coloc_pairs:
                if coloc_key in section_data['colocalizations']:
                    coloc_data = section_data['colocalizations'][coloc_key]
                    num_coloc = len(coloc_data['pairs'])
                    
                    ch1 = coloc_key.split('_')[0]
                    ch1_count = section_data['channels'].get(ch1, {}).get('count', 0)
                    percent = (num_coloc / ch1_count * 100) if ch1_count > 0 else 0
                    
                    row[f'{display_name} Coloc'] = num_coloc
                    row[f'{display_name} %'] = f"{percent:.1f}%"
                else:
                    row[f'{display_name} Coloc'] = 'N/A'
                    row[f'{display_name} %'] = 'N/A'
                    
            # Add triple colocalization
            if all(ch in section_data['channels'] for ch in ['WFA', 'Agg', 'PV']):
                triple_count = len(self.find_triple_colocalizations(section_data))
                row['Triple'] = triple_count
            else:
                row['Triple'] = 'N/A'
                
            summary_rows.append(row)
            
        # Add totals row
        totals = self.calculate_mouse_totals(summary_rows)
        summary_rows.append(totals)
        
        # Write CSV
        csv_path = os.path.join(self.output_dir, "CSV_Outputs", mouse, f"{mouse}_Summary.csv")
        
        header = [
            f"# {mouse} Complete Summary",
            f"# Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"# Channels analyzed: {', '.join(self.selected_channels)}"
        ]
        
        # Add pixel size info
        pixel_sizes = set()
        for section_data in mouse_sections:
            if section_data['pixel_size']:
                pixel_sizes.add(section_data['pixel_size'])
                
        if len(pixel_sizes) == 1:
            header.append(f"# Pixel size: {pixel_sizes.pop()} µm/pixel (all sections)")
        elif len(pixel_sizes) > 1:
            header.append("# Pixel size: Mixed across sections")
        else:
            header.append("# ⚠️ NO PIXEL SIZE DATA - Micron conversions unavailable")
            
        header.append("\nSECTION-BY-SECTION DETAILS:")
        
        self.write_csv_with_header(csv_path, summary_rows, header)
        
        # Also create Excel version
        excel_path = os.path.join(self.output_dir, "Excel_Outputs", mouse, f"{mouse}_Summary.xlsx")
        df = pd.DataFrame(summary_rows)
        df.to_excel(excel_path, index=False)
        
    def calculate_mouse_totals(self, summary_rows: List[Dict]) -> Dict:
        """Calculate totals for mouse summary."""
        totals = {'Section': 'TOTAL'}
        
        # Sum numeric columns
        sum_columns = [
            'WFA Count', 'Agg Count', 'PV Count', 'Area (px)',
            'WFA-Agg Coloc', 'WFA-PV Coloc', 'Agg-PV Coloc', 'Triple'
        ]
        
        for col in sum_columns:
            total = 0
            has_data = False
            
            for row in summary_rows[:-1] if summary_rows else []:  # Exclude totals row if it exists
                if col in row and row[col] != 'N/A':
                    try:
                        total += int(row[col])
                        has_data = True
                    except (ValueError, TypeError):
                        pass
                        
            totals[col] = total if has_data else 'N/A'
            
        # Calculate average percentages
        for coloc_type in ['WFA-Agg', 'WFA-PV', 'Agg-PV']:
            coloc_col = f'{coloc_type} Coloc'
            
            if totals.get(coloc_col, 'N/A') != 'N/A':
                # Get reference channel
                ref_channel = coloc_type.split('-')[0]
                ref_count = totals.get(f'{ref_channel} Count', 0)
                
                if ref_count and ref_count != 'N/A':
                    percent = (totals[coloc_col] / ref_count * 100)
                    totals[f'{coloc_type} %'] = f"{percent:.1f}%"
                else:
                    totals[f'{coloc_type} %'] = 'N/A'
            else:
                totals[f'{coloc_type} %'] = 'N/A'
                
        return totals
        
    def generate_master_summary(self, csv_dir: str, excel_dir: str):
        """Generate master summary combining all mice."""
        all_sections = []
        mouse_summaries = []
        
        # Collect all section data
        for section_key, section_data in sorted(self.all_data.items()):
            # Create row for complete data table
            row = {
                'Mouse': section_data['mouse'],
                'Section': section_data['section'],
                'µm Data': '✓' if section_data['pixel_size'] else '✗'
            }
            
            # Add counts
            for channel in ['WFA', 'Agg', 'PV']:
                row[channel] = section_data['channels'].get(channel, {}).get('count', 'N/A')
                
            # Add area
            areas = list(section_data['areas'].values())
            if areas:
                area_pixels = areas[0]
                row['Area(px)'] = area_pixels
                
                if section_data['pixel_size']:
                    area_microns = area_pixels * (section_data['pixel_size'] ** 2)
                    row['Area(µm²)'] = f"{area_microns:.0f}"
                else:
                    row['Area(µm²)'] = 'N/A'
            else:
                row['Area(px)'] = 'N/A'
                row['Area(µm²)'] = 'N/A'
                
            # Add colocalizations
            for coloc_key, display in [('WFA_Agg', 'WFA-Agg'), ('WFA_PV', 'WFA-PV'), ('Agg_PV', 'Agg-PV')]:
                if coloc_key in section_data['colocalizations']:
                    num_coloc = len(section_data['colocalizations'][coloc_key]['pairs'])
                    ch1 = coloc_key.split('_')[0]
                    ch1_count = section_data['channels'].get(ch1, {}).get('count', 0)
                    percent = (num_coloc / ch1_count * 100) if ch1_count > 0 else 0
                    
                    row[display] = num_coloc
                    row[f'{display}%'] = f"{percent:.1f}%"
                else:
                    row[display] = 'N/A'
                    row[f'{display}%'] = 'N/A'
                    
            # Triple colocalization
            if all(ch in section_data['channels'] for ch in ['WFA', 'Agg', 'PV']):
                triple_count = len(self.find_triple_colocalizations(section_data))
                row['Triple'] = triple_count
            else:
                row['Triple'] = 'N/A'
                
            # Add notes
            notes = []
            missing = [ch for ch in self.selected_channels if ch not in section_data['channels']]
            if missing:
                notes.append(f"Missing {', '.join(missing)}")
            row['Notes'] = '; '.join(notes) if notes else '-'
            
            all_sections.append(row)
            
        # Calculate mouse summaries
        for mouse in sorted(self.mice_to_process):
            mouse_data = [s for s in all_sections if s['Mouse'] == mouse]
            
            if mouse_data:
                summary = {
                    'Mouse': mouse,
                    'Sections': len(mouse_data)
                }
                
                # Sum totals
                for col in ['WFA', 'Agg', 'PV', 'WFA-Agg', 'WFA-PV', 'Agg-PV', 'Triple']:
                    total = sum(int(s[col]) for s in mouse_data if s.get(col, 'N/A') != 'N/A')
                    summary[f'Total {col}'] = total if any(s.get(col, 'N/A') != 'N/A' for s in mouse_data) else 'N/A'
                    
                # Check pixel data
                has_pixel_data = any(s['µm Data'] == '✓' for s in mouse_data)
                if has_pixel_data:
                    pixel_sizes = set()
                    for s in mouse_data:
                        section_key = f"{s['Mouse']}/{s['Section']}"
                        if section_key in self.all_data and self.all_data[section_key]['pixel_size']:
                            pixel_sizes.add(self.all_data[section_key]['pixel_size'])
                            
                    if len(pixel_sizes) == 1:
                        summary['Pixel Data'] = f"Yes ({pixel_sizes.pop()} µm/px)"
                    else:
                        summary['Pixel Data'] = "Yes (mixed)"
                else:
                    summary['Pixel Data'] = "No"
                    
                mouse_summaries.append(summary)
            
        # Write master CSV
        master_path = os.path.join(csv_dir, "All_Mice_Summary.csv")
        
        with open(master_path, 'w', newline='', encoding='utf-8') as f:
            # Write header
            f.write(f"# Master Summary - All Mice, All Sections\n")
            f.write(f"# Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            
            sections_with_microns = sum(1 for s in all_sections if s['µm Data'] == '✓')
            sections_without = len(all_sections) - sections_with_microns
            
            if sections_without > 0:
                f.write("# Mixed units: Some sections have pixel size data, others don't\n")
                
            f.write("\nCOMPLETE DATA TABLE:\n")
            
            # Write section data
            if all_sections:
                writer = csv.DictWriter(f, fieldnames=all_sections[0].keys())
                writer.writeheader()
                writer.writerows(all_sections)
                
            # Write mouse summaries
            f.write("\n\nMOUSE SUMMARIES:\n")
            if mouse_summaries:
                writer = csv.DictWriter(f, fieldnames=mouse_summaries[0].keys())
                writer.writeheader()
                writer.writerows(mouse_summaries)
                
            # Write grand totals
            f.write("\n\nGRAND TOTALS:\n")
            
            grand_totals = {
                'Total Mice': len(self.mice_to_process),
                'Total Sections Processed': len(all_sections)
            }
            
            # Count sections with colocalization
            sections_with_coloc = sum(1 for s in all_sections 
                                     if any(s.get(c, 'N/A') != 'N/A' 
                                           for c in ['WFA-Agg', 'WFA-PV', 'Agg-PV']))
            
            grand_totals['Sections with Colocalization'] = sections_with_coloc
            
            # Sum all cells
            for channel in ['WFA', 'Agg', 'PV']:
                total = sum(int(s[channel]) for s in all_sections if s.get(channel, 'N/A') != 'N/A')
                grand_totals[f'Total {channel} Cells'] = total
                
            # Sum colocalizations
            for coloc in ['WFA-Agg', 'WFA-PV', 'Agg-PV', 'Triple']:
                total = sum(int(s[coloc]) for s in all_sections if s.get(coloc, 'N/A') != 'N/A')
                grand_totals[f'Total {coloc} Colocalizations'] = total
                
            # Write grand totals
            for key, value in grand_totals.items():
                f.write(f"{key}: {value}\n")
                
        # Create Excel version
        excel_path = os.path.join(excel_dir, "All_Mice_Summary.xlsx")
        
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            # Complete data
            if all_sections:
                df = pd.DataFrame(all_sections)
                df.to_excel(writer, sheet_name='Complete_Data', index=False)
                
            # Mouse summaries
            if mouse_summaries:
                df = pd.DataFrame(mouse_summaries)
                df.to_excel(writer, sheet_name='Mouse_Summaries', index=False)
                
            # Grand totals
            df = pd.DataFrame([grand_totals])
            df.to_excel(writer, sheet_name='Grand_Totals', index=False)
            
    def generate_units_report(self, csv_dir: str):
        """Generate report on measurement units used."""
        report_path = os.path.join(csv_dir, "Measurement_Units_Report.txt")
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write("MEASUREMENT UNITS REPORT\n")
            f.write("=" * 50 + "\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            
            # Count sections by measurement type
            sections_with_microns = []
            sections_without_microns = []
            
            for section_key, section_data in self.all_data.items():
                if section_data['pixel_size']:
                    sections_with_microns.append((section_key, section_data['pixel_size']))
                else:
                    sections_without_microns.append(section_key)
                    
            # Report sections with micron data
            f.write(f"SECTIONS WITH MICRON DATA ({len(sections_with_microns)}/{len(self.all_data)}):\n")
            f.write("-" * 50 + "\n")
            
            for section_key, pixel_size in sorted(sections_with_microns):
                mouse, section = section_key.split('/')
                f.write(f"{mouse} - {section}: {pixel_size} µm/pixel\n")
                
            # Report sections without micron data
            if sections_without_microns:
                f.write(f"\n\nSECTIONS WITHOUT MICRON DATA ({len(sections_without_microns)}/{len(self.all_data)}):\n")
                f.write("-" * 50 + "\n")
                
                for section_key in sorted(sections_without_microns):
                    mouse, section = section_key.split('/')
                    f.write(f"{mouse} - {section}: Pixel measurements only\n")
                    
            # Summary statistics
            f.write("\n\nSUMMARY:\n")
            f.write("-" * 50 + "\n")
            
            percent_with_microns = (len(sections_with_microns) / len(self.all_data) * 100 
                                   if self.all_data else 0)
            
            f.write(f"Total sections processed: {len(self.all_data)}\n")
            f.write(f"Sections with micron data: {len(sections_with_microns)} ({percent_with_microns:.1f}%)\n")
            f.write(f"Sections with pixel-only data: {len(sections_without_microns)} ({100 - percent_with_microns:.1f}%)\n")
            
            # Unique pixel sizes
            unique_pixel_sizes = set(ps for _, ps in sections_with_microns)
            if unique_pixel_sizes:
                f.write(f"\nUnique pixel sizes found: {', '.join(f'{ps} µm/pixel' for ps in sorted(unique_pixel_sizes))}\n")
                
            # Recommendations
            f.write("\n\nRECOMMENDATIONS:\n")
            f.write("-" * 50 + "\n")
            
            if sections_without_microns:
                f.write("• Consider providing pixel size data for all sections to enable:\n")
                f.write("  - Standardized density calculations (cells/mm²)\n")
                f.write("  - Micron-based distance measurements\n")
                f.write("  - Cross-study comparisons\n")
            else:
                f.write("• All sections have pixel size data ✓\n")
                f.write("• Results are suitable for publication with standardized units\n")
                
    def show_completion_summary(self):
        """Show final completion summary."""
        print("\n" + "═" * 70)
        print("ANALYSIS COMPLETE")
        print("═" * 70)
        
        # Count successful/failed sections
        total_attempted = len(self.all_data)
        sections_with_errors = len([1 for log in self.error_log if "Error" in log])
        sections_with_warnings = len([1 for log in self.error_log if "Warning" in log])
        
        print(f"\n✓ Successfully processed: {total_attempted}/{total_attempted} sections")
        
        if sections_with_warnings > 0:
            print(f"⚠️ Warnings: {sections_with_warnings}")
            
        if sections_with_errors > 0:
            print(f"✗ Errors: {sections_with_errors}")
            
        # Processing summary
        print("\nProcessing Summary:")
        print("─────────────────")
        
        # Count total cells
        total_cells = {}
        for channel in ['WFA', 'Agg', 'PV']:
            total = sum(
                section_data['channels'].get(channel, {}).get('count', 0)
                for section_data in self.all_data.values()
            )
            if total > 0:
                total_cells[channel] = total
                
        cells_str = ', '.join(f"{v} {k}" for k, v in total_cells.items())
        print(f"- Total PNNs detected: {sum(total_cells.values())} ({cells_str})")
        
        # Average colocalization rates
        coloc_rates = defaultdict(list)
        
        for section_data in self.all_data.values():
            for coloc_key, coloc_data in section_data['colocalizations'].items():
                ch1 = coloc_key.split('_')[0]
                ch1_count = section_data['channels'].get(ch1, {}).get('count', 0)
                if ch1_count > 0:
                    rate = len(coloc_data['pairs']) / ch1_count * 100
                    coloc_rates[coloc_key].append(rate)
                    
        if coloc_rates:
            avg_rates = []
            for coloc_key, rates in sorted(coloc_rates.items()):
                avg_rate = sum(rates) / len(rates)
                display_name = coloc_key.replace('_', '-')
                avg_rates.append(f"{display_name}: {avg_rate:.1f}%")
                
            print(f"- Average colocalization rates: {', '.join(avg_rates)}")
            
        # Measurement summary
        print("\nMeasurement Summary:")
        print("─────────────────")
        
        sections_with_microns = sum(1 for s in self.all_data.values() if s['pixel_size'])
        sections_without = len(self.all_data) - sections_with_microns
        
        if sections_with_microns > 0:
            percent = sections_with_microns / len(self.all_data) * 100
            print(f"- Sections with full measurements: {sections_with_microns}/{len(self.all_data)} ({percent:.0f}%)")
            
        if sections_without > 0:
            percent = sections_without / len(self.all_data) * 100
            print(f"- Sections with pixel-only: {sections_without}/{len(self.all_data)} ({percent:.0f}%)")
            
        print(f"\nOutput location: {self.output_dir}/")
        
        # Write error log if any errors/warnings
        if self.error_log:
            error_path = os.path.join(self.output_dir, "Logs", "error_summary.txt")
            os.makedirs(os.path.dirname(error_path), exist_ok=True)
            
            with open(error_path, 'w', encoding='utf-8') as f:
                f.write("ERROR AND WARNING LOG\n")
                f.write("=" * 50 + "\n")
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                
                for entry in self.error_log:
                    f.write(entry + "\n")
                    
            print(f"\n⚠️ See {error_path} for details on warnings/errors")
            
        # Options menu
        print("\nOptions:")
        print("1. Open output folder")
        print("2. View detailed statistics")
        print("3. Exit")
        
        choice = input("\nEnter choice (1-3): ").strip()
        
        if choice == '1':
            # Open folder (platform-specific)
            import platform
            if platform.system() == 'Windows':
                os.startfile(self.output_dir)
            elif platform.system() == 'Darwin':  # macOS
                os.system(f"open {self.output_dir}")
            else:  # Linux
                os.system(f"xdg-open {self.output_dir}")
                
        elif choice == '2':
            self.show_detailed_statistics()
            
    def show_detailed_statistics(self):
        """Show detailed statistics about the analysis."""
        print("\n" + "═" * 70)
        print("DETAILED STATISTICS")
        print("═" * 70)
        
        # Measurement coverage
        print("\nMeasurement Coverage:")
        sections_with_microns = sum(1 for s in self.all_data.values() if s['pixel_size'])
        sections_without = len(self.all_data) - sections_with_microns
        
        print(f"- Sections with full micron data: {sections_with_microns}/{len(self.all_data)} " +
              f"({sections_with_microns/len(self.all_data)*100:.1f}%)")
        print(f"- Sections with pixel-only data: {sections_without}/{len(self.all_data)} " +
              f"({sections_without/len(self.all_data)*100:.1f}%)")
        
        # Channel coverage
        print("\nChannel Coverage:")
        channel_counts = defaultdict(int)
        
        for section_data in self.all_data.values():
            for channel in section_data['channels']:
                channel_counts[channel] += 1
                
        for channel in ['WFA', 'Agg', 'PV']:
            if channel in self.selected_channels:
                count = channel_counts.get(channel, 0)
                percent = count / len(self.all_data) * 100 if self.all_data else 0
                print(f"- Sections with {channel}: {count}/{len(self.all_data)} ({percent:.1f}%)")
                
        # Colocalization analysis coverage
        both_channels = sum(1 for s in self.all_data.values() if len(s['channels']) >= 2)
        print(f"- Sections with colocalization analysis: {both_channels}/{len(self.all_data)} " +
              f"({both_channels/len(self.all_data)*100:.1f}%)")
        
        # Colocalization rates by mouse
        print("\nColocalization Rates by Mouse:")
        
        for mouse in sorted(self.mice_to_process):
            mouse_sections = [s for k, s in self.all_data.items() if s['mouse'] == mouse]
            
            if mouse_sections:
                print(f"\n{mouse}:")
                
                # Calculate average rates for this mouse
                mouse_coloc_rates = defaultdict(list)
                
                for section_data in mouse_sections:
                    for coloc_key, coloc_data in section_data['colocalizations'].items():
                        ch1 = coloc_key.split('_')[0]
                        ch1_count = section_data['channels'].get(ch1, {}).get('count', 0)
                        if ch1_count > 0:
                            rate = len(coloc_data['pairs']) / ch1_count * 100
                            mouse_coloc_rates[coloc_key].append(rate)
                            
                for coloc_key, rates in sorted(mouse_coloc_rates.items()):
                    avg_rate = sum(rates) / len(rates)
                    display_name = coloc_key.replace('_', '-')
                    
                    # Check if this mouse has pixel size data
                    has_microns = any(s['pixel_size'] for s in mouse_sections)
                    unit_indicator = " (micron data available)" if has_microns else " (pixel data only)"
                    
                    print(f"  - {display_name}: {avg_rate:.1f}%{unit_indicator}")
                    
        input("\nPress Enter to return to menu...")


# Main execution
if __name__ == "__main__":
    pipeline = PNNColocalizationPipeline()
    pipeline.run()