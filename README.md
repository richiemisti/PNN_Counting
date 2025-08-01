PNN_Counting: A Deep Learning Pipeline for Perineuronal Net and Interneuron Quantification
===========================================================================================

Overview
--------
This repository provides an adapted and streamlined version of the deep learning pipeline originally developed by Lupori et al. (Cell Reports, 2023) for quantifying perineuronal nets (PNNs) and parvalbumin + (PV+) neurons in brain tissue. This modified version has been tailored for confocal image analysis in the context of hypertension research, specifically using the DOCA-salt mouse model in the Santisteban Lab at Vanderbilt University Medical Center.

Our implementation supports automated detection, colocalization, and intensity quantification of PNNs labeled with WFA and aggrecan, as well as PV⁺ interneurons, in hippocampal brain sections.

Key Features
------------
- Automated organization of raw confocal image data into the appropriate folder structure
- Inference using pretrained Faster R-CNN and ordinal regression models for structure detection and scoring
- Visualization of model predictions overlaid on the original images
- Colocalization analysis across WFA, aggrecan, and PV channels
- Quantification of WFA intensity at each detected PNN location
- Structured output of results in CSV and Excel formats

Modifications from the Original Repository
------------------------------------------
This work builds upon the original repository: https://github.com/ciampluca/counting_perineuronal_nets

The following modifications were made:
- **Added**:
  - `structure.py`: Organizes raw image data into the expected directory format
  - `run_pipeline.py`: Wrapper script to automate prediction and visualization generation
  - `colocalization.py`: Detects spatial overlaps between PNN and PV localizations with optional image and statistical outputs
  - `Calculate Intensities of PNNs_Clean.ipynb`: Jupyter notebook for grayscale intensity quantification
  - `Mice_WFA/` and `PV_Mice/`: Example directory scaffolds for user reference

This adapted version focuses exclusively on inference and analysis workflows and is optimized for use with pre-segmented, masked confocal images.

Repository
----------
The pipeline (WITHOUT DEEP LEARNING MODELS) is hosted here: https://github.com/richiemisti/PNN_Counting.git
- Download DL Models here (see below for specifics): https://github.com/ciampluca/counting_perineuronal_nets (navigate to their README.md file to find their trained models)

Contributors and Acknowledgements
---------------------------------
Original Authors (Lupori et al., 2023):
- Luca Ciampi, Leonardo Lupori, Tommaso Pizzorusso, et al.

Santisteban Lab Contributors:
- **Richie Mistichelli** – Lead contributor for `structure.py`, `run_pipeline.py`, `colocalization.py`, and documentation. Pipeline adaptation and integration.
- **Ismary Blanco** – Scientific guidance, conceptual contributions, and technical feedback.
- **Florencia Martinez Addiego** – Developed `Calculate Intensities of PNNs_Clean.ipynb`, `draw_and_measure.py`, and `rolling_ball.py` for downstream quantification.

Please refer to the `NOTICE` and `CONTRIBUTORS.md` files for full attribution.

License
-------
This project remains under the terms of the original **Apache License 2.0**, as provided in the Lupori et al. repository. All modifications and new contributions are also distributed under the same license. See `LICENSE` for full details.

Getting Started: Setup Instructions
-------

IMPORTANT: Download the Deep Learning Models
-------------------------------------
There are **four total models** used in the pipeline (two for PNNs, two for PV):
- PNN Localization – `FRCNN-640`
- PNN Scoring – `Ordinal Regression`
- PV Localization – `FRCNN-640`
- PV Scoring – `Ordinal Regression`

They can be downloaded from: https://github.com/ciampluca/counting_perineuronal_nets
- navigate to their README.md file to find them

IMPORTANT: Folder Naming Rules
-------------------------------------
- Use `Mice_WFA/` or `Mice_Agg/` as parent folders for **WFA or Aggrecan images**
- Use `PV_Mice/` for **PV images**
- These names MUST be exact or the pipeline will skip them.
- Each mouse should be in its own folder (e.g., `Mouse_4371-4`) and contain slice images (`1L.tif`, `1R.tif`, etc.)


Environment Setup (One Time)
================

1. Install Miniconda:  
   https://www.anaconda.com/docs/getting-started/miniconda/main  
   - Download Python 3.8 or later (Windows installer)
   - During installation, check:
     - "Add Miniconda to my PATH"
     - "Register Miniconda as system Python"

2. Verify install:
   ```
   conda --version
   ```

3. Clone this repository:
   ```
   git clone https://github.com/richiemisti/PNN_Counting.git
   ```

4. Install VS Code:
   https://code.visualstudio.com/

   - Open the repo folder (`PNN_Counting`)
   - VS Code will auto-prompt to install Python extension – accept it

5. Open Terminal (CMD preferred, not PowerShell) and run:
   ```
   conda create -n peri python=3.8
   conda activate peri
   ```

6. Install PyTorch (CPU-only):
   ```
   pip install torch==1.11.0+cpu torchvision==0.12.0+cpu --extra-index-url https://download.pytorch.org/whl/cpu
   ```

7. Install remaining dependencies:
   ```
   pip install -r requirements.txt
   ```


Image Preprocessing (ImageJ)
=========

Use ImageJ to:
- Convert Z-stacks → SUM projections
- Convert to 8-bit grayscale
- Rename based on brain region (automated via macro or manual)
- Ensure consistent naming: `Mouse_XXXX/1L.tif`, `2L.tif`, etc.
- Names must match across channels for proper analysis

```javascript
// === start of macro ===
// This ImageJ macro automates preprocessing of .czi confocal image files. It allows the user to select a parent folder and then either one or all subfolders to process. For each valid .czi file that includes a left or right hemisphere designation, it performs a SUM Z-projection, applies grayscale and 8-bit conversion, and then saves two TIFF images with standardized names for hippocampus and CA1 regions. This prepares the data for downstream PNN and PV detection workflows.
// Choose parent folder, optionally pick ONE subfolder or ALL.
// For each valid .czi: SUM Z, Grays LUT, convert to 8-bit, save two TIFFs with "_SUM" and trailing underscore.

root = getDirectory("Choose PARENT folder containing mouse subfolders");
print("ROOT folder: " + root);

// Build list of subfolders
entries = getFileList(root);
count = 0;
for (i=0; i<entries.length; i++) {
    if (File.isDirectory(root + entries[i])) count++;
}
dirs = newArray(count);
idx = 0;
for (i=0; i<entries.length; i++) {
    f = entries[i];
    if (File.isDirectory(root + f)) {
        dirs[idx] = f;
        idx++;
    }
}

// Add "__ALL__" option
choiceList = newArray(dirs.length + 1);
choiceList[0] = "__ALL__";
for (i=0; i<dirs.length; i++) choiceList[i+1] = dirs[i];

// Ask user which to process
Dialog.create("Choose subfolder");
Dialog.addChoice("Process which folder?", choiceList, choiceList[0]);
Dialog.show();
target = Dialog.getChoice();
print("Selected: " + target);

// Process
if (target == "__ALL__") {
    for (i=0; i<dirs.length; i++) processFolder(root + dirs[i] + File.separator);
} else {
    processFolder(root + target + File.separator);
}

print("\nAll done!");
// === end of macro ===


// ---------- helper function ----------
function processFolder(mouseDir) {
    print("\nProcessing folder: " + mouseDir);
    files = getFileList(mouseDir);

    for (i = 0; i < files.length; i++) {
        name = files[i];

        // Filters
        if (!endsWith(name, ".czi"))        continue;
        if (indexOf(name, "section_") < 0)  continue;
        hasR = indexOf(name, "_right_") > -1;
        hasL = indexOf(name, "_left_")  > -1;
        if (!hasR && !hasL)                 continue;

        // Mouse ID
        idStart = indexOf(name, "IB");
        if (idStart < 0) { print("  ⚠ No 'IB' in: " + name); continue; }
        idEnd   = indexOf(name, "_", idStart);
        mouseID = substring(name, idStart, idEnd);

        // Section #
        secKey     = "section_";
        secPos     = indexOf(name, secKey) + lengthOf(secKey);
        secEnd     = indexOf(name, "_", secPos);
        sectionNum = substring(name, secPos, secEnd);

        // Side
        if (hasR) sideCode = "R"; else sideCode = "L";

        // Outputs (note trailing underscore)
        hippOut = mouseID + "_SUM_HIPP_" + sectionNum + sideCode + "_" + ".tif";
        ca1Out  = mouseID + "_SUM_CA1_"  + sectionNum + sideCode + "_" + ".tif";

        fullPath = mouseDir + name;
        print("→ " + name + "  →  " + hippOut + " / " + ca1Out);

        // Open
        run("Bio-Formats Importer", "open=[" + fullPath + "] autoscale color_mode=Default view=Hyperstack stack_order=XYCZT");
        origTitle = getTitle();

        // Z SUM
        run("Z Project...", "projection=[Sum Slices]");
        run("Grays");
        run("8-bit");

        // Save
        saveAs("Tiff", mouseDir + hippOut);
        saveAs("Tiff", mouseDir + ca1Out);

        // Close images
        close(); // projection
        selectWindow(origTitle);
        close();
    }
}


Folder Restructuring
==============================

1. Inside `PNN_Counting/`, make your parent folders:
   ```
   Mice_WFA/
   PV_Mice/
   Mice_Agg/ (optional)
   ```

2. Inside each parent, add mouse folders:
   ```
   Mice_WFA/Mouse_IB60/
   PV_Mice/Mouse_IB60/
   ```

3. Place your processed `.tif` images in the corresponding mouse folders.

Example Structure: 

PNN_Counting/
├── Mice_WFA/
│   └── Mouse_IB60/
│       ├── IB60_SUM_CA1_1L_WFA.tif
│       ├── IB60_SUM_CA1_1R_WFA.tif
│       └── ...
├── Mice_Agg/
│   └── Mouse_IB60/
│       ├── IB60_SUM_CA1_1L_AGG.tif
│       ├── IB60_SUM_CA1_1R_AGG.tif
│       └── ...
├── PV_Mice/
│   └── Mouse_IB60/
│       ├── IB60_SUM_CA1_1L_PV.tif
│       ├── IB60_SUM_CA1_1R_PV.tif
│       └── ...

Tip: All channels (WFA, Agg, PV) must have identical mouse folder names and image section names (e.g., IB60_SUM_CA1_1L_XXX.tif) to ensure colocalization runs correctly.

4. Enter VS code ensuring that the PNN_Counting repo folder is open (click two pages icon in the top left to open folder)

5. Activate the conda environment if not already:
   ```
   conda activate peri

Tip: You must always be in this environment to get the code to work, not in (base)

4. Run the folder organizer after ensuring proper pre-processing structure (see step 3):
   ```
   python structure.py
   ```
This script will:
-Create a folder for each section (e.g., IB60_SUM_CA1_1L_WFA)
-Move the corresponding .tif image inside
-Prepare the structure for predictions

PNN_Counting/
├── Mice_WFA/
│   └── Mouse_IB60/
│       └── IB60_SUM_CA1_1L_WFA/
│           └── IB60_SUM_CA1_1L_WFA.tif
├── Mice_Agg/
│   └── Mouse_IB60/
│       └── IB60_SUM_CA1_1L_AGG/
│           └── IB60_SUM_CA1_1L_AGG.tif
├── PV_Mice/
│   └── Mouse_IB60/
│       └── IB60_SUM_CA1_1L_PV/
│           └── IB60_SUM_CA1_1L_PV.tif

Running the Pipeline
==============================

After structuring and masking are complete:

1. Activate the conda environment if not already:
   ```
   conda activate peri
   ```

2. Finalize folder structure if not already:
   ```
   python structure.py
   ```

3. Run predictions:
   ```
   python run_pipeline.py
   ```
run_pipeline.py Output per Section Folder:
- Original image (already there): *.tif
- Localization CSV file: localizations_<section_name>.csv
  - Contains predicted X, Y coordinates and bounding box metadata for each detected object (e.g., PNN, PV)
- Prediction images in a subfolder: <section_name>_predictions/
  - Contains 6 PNGs with bounding boxes overlaid on grayscale versions of the input

4. Colocalization (see below for more details, this is a complex script):
   ```
   python colocalization.py
   ```

   - Generates raw match files, plots, distance histograms, summary stats

Colocalization with colocalization.py
==============================================

Overview
--------
The `colocalization.py` script performs spatial colocalization analysis between detected structures across different staining channels. Specifically, it compares perineuronal net (PNN) markers (e.g., WFA, aggrecan) and interneuron markers (e.g., parvalbumin, PV) to determine overlap based on user-defined proximity thresholds.

This step is critical for quantifying structural relationships between ECM components and PV⁺ interneurons in the hippocampus.

Pre-requisites
--------------
Before running the script, ensure the following:
- Deep learning predictions have been generated using `run_pipeline.py` for all relevant channels (WFA, Agg, PV).
- Folder and image names are consistent across channels. For example:
  ├── Mice_WFA/Mouse_XXXX/IB60_SUM_CA1_1L_WFA/
  ├── Mice_Agg/Mouse_XXXX/IB60_SUM_CA1_1L_AGG/
  ├── PV_Mice/Mouse_XXXX/IB60_SUM_CA1_1L_PV/

- Each section folder must contain:
  - The original `.tif` image
  - A CSV of predicted coordinates (e.g., `localizations_IB60_SUM_CA1_1L_WFA.csv`)

Running the Script
------------------
1. Activate the Conda environment:
   ```
   conda activate peri
   ```

2. Execute the colocalization analysis:
   ```
   python colocalization.py
   ```

3. You will be prompted to:
   - Enter a pixel distance threshold (e.g., 10 pixels)
   - Choose whether to perform WFA↔PV, WFA↔Agg, Agg↔PV, or full triple-channel colocalization
   - Select whether to generate visual outputs and save unmatched points

Output Files
------------
For each mouse and section, the script will generate:
- Colocalized CSVs (e.g., `WFA_PV_colocalized.csv`)
- Optional unmatched points CSVs
- Image visualizations of matched points (if selected)
- NOTE: The most important output file is the colocalization CSV, which contains the indices of matched points along with their corresponding original indices from the individual channel localization CSVs.
  - You can find these files at a path similar to: CSV_Outputs/Mouse_IB60/IB60_SUM_CA1_1L/
	- Inside that folder, look for files named: 6_colocalization_*.csv
	- These files summarize colocalized coordinates, distances between matched points, and the mapping back to the original detections.

- A summary Excel file:
  - `colocalization_summary.xlsx`

This Excel file includes:
- Counts of matched points per pair
- Percentages of colocalized points relative to total detected structures

Best Practices
--------------
- Use identical naming and coordinate space across images for accurate matching.
- Run the script once all channels have been processed.
- Avoid re-running colocalization if outputs already exist unless changes are made.

For any issues or clarification, please contact the pipeline maintainer

Interactive Walkthrough for Colocalization.py (Recommended Inputs)
==================================================================

This guide outlines how to interact with `colocalization.py` using recommended responses based on standard Santisteban Lab use cases for DOCA-salt hypertension studies.

STEP 1: Channel Selection
-------------------------
Prompt:
> Which channels would you like to analyze?

Recommended Input:
```
all
```
Includes all possible pairwise and triple-channel colocalizations (if three channels present).

STEP 2: Mouse Selection
-----------------------
Prompt:
> Which mice would you like to process?

Recommended Input:
```
1
```
Selects all detected mice for analysis.

STEP 3: Pixel Size Configuration
--------------------------------
Prompt:
> How would you like to specify pixel measurements?

Recommended Input:
```
1
```
Uses pixel units only (no micron conversion). This is preferred due to inconsistent TIF metadata.

STEP 4: Colocalization Threshold
--------------------------------
Prompt:
> Enter colocalization threshold in PIXELS

Recommended Input:
```
60
```
This threshold works well with typical imaging resolution and PNN/PV proximity.

STEP 5: Additional Analysis Options
-----------------------------------
Prompt:
> Generate distance distribution reports? (y/n)

Recommended Input:
```
n
```
Skip unless specifically required.

Prompt:
> Generate visual overlays? (y/n)

Recommended Input:
```
y
```
Enables circle overlays on detected matches.

Follow-up Inputs:
- Circle thickness: `2`
- Detection diameter: `35`
- Composite thickness: `3`
- Color scheme (default):
  - WFA = Yellow
  - PV = Blue
  - Agg = Red
  - 2-way colocalized = White
  - 3-way colocalized = Magenta

Threshold Visualizations
------------------------
Prompt:
> Generate threshold-focused visualizations? (y/n)

Recommended Input:
```
y
```
- these images are the ones we referred to the most and are very important. I would recommend printing these for colocalizaiton visualization. 

- Circle diameter: `35`
- Visualization type: `2` (color-coded method)

Composite Background Settings
-----------------------------
Prompt:
> WFA contribution (0.1–0.9): 
```
.4
```

Prompt:
> PV contribution:
```
.6
```

This blend balances visibility across both markers.

Final Summary
-------------
Once configured, the script will display a full summary and begin processing. Outputs are saved to:
```
PNN_Counting/Analysis_<DATE>_<TIME>/
```

Includes:
- Raw colocalization CSVs
- Excel summaries
- Image overlays (if enabled)
- Run metadata and logs

PNN Intensity Quantification - Credit: Florencia Martinez Addiego
==============================

1. Developed by Florencia Martinez Addiego, open:
   ```
   Calculate Intensities of PNNs_Clean.ipynb
   ```


2. Run cells top to bottom using VS Code Jupyter notebook support.

3. Make sure:
   - You're using the **`peri`** conda environment
   - You've installed Jupyter in that environment:
     ```
     conda install jupyter
     ```

4. In the notebook, update:
   ```python
   parent_dir = '/Users/you/Documents/.../PNN_Counting'
   ```

   Right-click the repo folder → Copy path → Paste as string in the cell

5. Run remaining cells — intensity and binned results will be saved.

Questions?
==============================

Please contact:
- Richie Mistichelli — [rmistich@nd.edu]
- Ismary Blanco — [ismary.blanco@vumc.org]
