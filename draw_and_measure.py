import numpy as np
import pandas as pd
import cv2
import matplotlib.pyplot as plt
from PIL import Image, ImageDraw
from skimage import io, transform
from skimage.io import imread
from matplotlib import cm
cmap = cm.get_cmap('viridis')
from scipy.ndimage import binary_dilation
import os

def draw_and_measure(image, data_path, diameter, cmap):
    """
    Quality control step to draw a circle around PNNs that were counted and visualize them on the rolling ball radius output image

    Inputs:
    - image: image background subtracted - np.ndarray (H, W) # keep this grayscaled
    - data: path to csv file with columns ['Y', 'X', 'score']
    - diameter: diameter in pixels # the function from the PNN counter calls the diameter the radius
    - cmap: function mapping score to RGB tuple (0–1 scale)

    Outputs:
    - the image as an array
    - the mean intensity from each of the predicted PNNs
    """
    image = Image.fromarray(image)
    pil_draw = ImageDraw.Draw(image)
    image_np = np.array(image.convert('L'))  # Convert to grayscale for intensity
    # Initialize combined PNN mask once
    combined_pnn_mask = np.zeros(image_np.shape, dtype=bool)

    data = pd.read_csv(data_path)

    data_name = os.path.basename(data_path) # should give the name.csv file
    data_name2 = os.path.splitext(data_name)[0] 
    data_name_split = data_name2.split('_')
    mouse_name = data_name_split[1]
    mouse_staining = data_name_split[-1]

    section = data_name_split[3]+'_'+data_name_split[4]
    
    means = []
    stdev = []
    sterr = []

    for r, c, s in data[['Y', 'X', 'score']].values:
        # Set color
        color = int(255 * cmap(s)[0])


        # Bounding box for circle
        y0, x0 = int(r - diameter / 2), int(c - diameter / 2)
        y1, x1 = int(r + diameter / 2), int(c + diameter / 2)

        # Draw circle
        pil_draw.ellipse([x0, y0, x1, y1], outline=color, width=3)

        # Create mask for the current circle
        mask = Image.new("L", image_np.shape[::-1], 0)  # (width, height)
        draw_mask = ImageDraw.Draw(mask)
        draw_mask.ellipse([x0, y0, x1, y1], fill=1)
        mask_np = np.array(mask).astype(bool)
        
        combined_pnn_mask |= mask_np

        # Compute mean intensity
        masked_image = np.where(mask_np, image_np, np.nan)
        mean_val = np.nanmean(masked_image)
        means.append(mean_val)
        stdev.append(np.nanstd(masked_image))
        sterr.append(np.nanstd(masked_image)/np.sqrt(np.size(np.where(~np.isnan(masked_image)))))

    df = pd.DataFrame(np.column_stack([means,sterr,stdev]),columns =['Mean Intensity','Sterr Intensity','Stdev Intensity'])
    df.insert(loc=0,column='Staining',value = mouse_staining)
    df.insert(loc=0,column='Section',value = section)
    df.insert(loc=0,column='Mouse Name',value = mouse_name)

    avg_means = np.nanmean(means)
    avg_stdev = np.nanstd(means)
    avg_sterr = np.nanstd(means)/np.sqrt(np.size(means))
    df_master = pd.DataFrame(np.column_stack([avg_means,avg_sterr,avg_stdev]),columns =['Avg Mean Intensity','Avg Sterr Intensity','Avg Stdev Intensity'])
    df_master.insert(loc=0,column='Staining',value = mouse_staining)
    df_master.insert(loc=0,column='Section',value = section)
    df_master.insert(loc=0,column='Mouse Name',value = mouse_name)

    return df, df_master, np.array(image), means, stdev, sterr, combined_pnn_mask

# example usage df, df_master output_image, intensity_means, intesnity_stdev, intensity_sterr, combined_pnn_mask = draw_and_measure(img_clean_gray, loc_csv, radius=40, cmap=cmap)