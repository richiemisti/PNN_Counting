import cv2
from skimage.restoration import rolling_ball

def rolling_ball_radius(image_path,radius):
    """
    Rolling ball radius function does the rolling ball subtraction 
    Inputs:
    - image: np.ndarray (H, W), keep grayscaled
    - radius: radius in pixels to smooth over 
    
    Notes: For this method to give meaningful results, the radius of the ball (or typical size of the kernel, in the general case) should be larger than the typical size of the image features of interest

    Returns:
    - img_bg_subtracted: np.ndarray, image with background subtracted
    """
    image = cv2.imread(image_path,cv2.IMREAD_GRAYSCALE)
    bg = rolling_ball(image,radius=radius)
    img_bg_subtracted = image - bg

    cv2.imwrite('img_rolling_ball_applied_r' + str(radius)+ '.jpg',img_bg_subtracted)

    return img_bg_subtracted

#example img_bg_subtracted = rolling_ball_radius(image_path,radius=25)