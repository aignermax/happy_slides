import os
from pptx import Presentation
from pptx.util import Inches
from natsort import natsorted
from PIL import Image, ExifTags
import io

# Description:
# This script processes all images located in the script's directory and its subdirectories.
# It creates a PowerPoint presentation where each image is placed on a separate slide.
# Images are correctly oriented based on their EXIF data and are inserted in the order of their folders and filenames.

# Path to the main folder containing images
main_folder = r"."

# Set the working directory
os.chdir(main_folder)
print(f"Main folder set to: {main_folder}")
print(f"Working directory set to: {os.getcwd()}")

# Create a PowerPoint presentation object
presentation = Presentation()

# Function to read EXIF data and correct image orientation
def fix_image_orientation(image_path):
    try:
        image = Image.open(image_path)
        for orientation in ExifTags.TAGS.keys():
            if ExifTags.TAGS[orientation] == 'Orientation':
                break
        exif = image._getexif()
        if exif is not None:
            orientation = exif.get(orientation, 1)
            if orientation == 3:  # Rotate 180 degrees
                image = image.rotate(180, expand=True)
            elif orientation == 6:  # Rotate 270 degrees
                image = image.rotate(270, expand=True)
            elif orientation == 8:  # Rotate 90 degrees
                image = image.rotate(90, expand=True)
        # Convert the image to RGB mode to save as JPEG
        image = image.convert('RGB')
        return image
    except Exception as e:
        print(f"Error processing {image_path}: {e}")
        return None

# Function to add an image with black borders if necessary
def add_image_with_black_borders(slide, image):
    # Convert slide dimensions from EMUs to pixels
    dpi = 96  # Typical DPI for PowerPoint slides
    slide_width_px = int(presentation.slide_width / 914400 * dpi)
    slide_height_px = int(presentation.slide_height / 914400 * dpi)

    # Calculate scaling factors for width and height
    scale_w = slide_width_px / image.width
    scale_h = slide_height_px / image.height

    # Use the smaller scaling factor to maintain the image's aspect ratio
    scale = min(scale_w, scale_h)

    # New image dimensions after scaling
    new_width = int(image.width * scale)
    new_height = int(image.height * scale)

    # Create a black background image with slide dimensions in pixels
    black_background = Image.new("RGB", (slide_width_px, slide_height_px), "black")

    # Resize the image to fit within the slide
    resized_image = image.resize((new_width, new_height), Image.ANTIALIAS)

    # Paste the resized image onto the black background, centered
    left = (slide_width_px - new_width) // 2
    top = (slide_height_px - new_height) // 2
    black_background.paste(resized_image, (left, top))

    # Save the result to an in-memory byte stream
    image_stream = io.BytesIO()
    black_background.save(image_stream, format="JPEG")
    image_stream.seek(0)

    # Add the black background with the image to the slide
    slide.shapes.add_picture(image_stream, Inches(0), Inches(0), width=presentation.slide_width, height=presentation.slide_height)


# Function to collect and sort images from a folder (including subfolders)
def get_images_sorted(folder):
    images = []
    for root, _, files in os.walk(folder):
        for file in files:
            if file.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif')):
                images.append(os.path.join(root, file))
    # Sort images alphabetically
    return natsorted(images)

# Collect all images (including top-level folder and subfolders)
all_images = get_images_sorted(main_folder)

# Process images in sorted order
total_images = len(all_images)
processed_images = 0

for image_path in all_images:
    # Load the image with the correct orientation
    fixed_image = fix_image_orientation(image_path)
    if fixed_image is None:
        continue

    # Add a new blank slide
    slide = presentation.slides.add_slide(presentation.slide_layouts[5])  # Blank slide layout

    # Add the image with black borders to the slide
    add_image_with_black_borders(slide, fixed_image)

    processed_images += 1
    print(f"Processed {processed_images}/{total_images} images: {image_path}")

# Save the presentation
output_file = "Photos_Presentation.pptx"
presentation.save(output_file)
print(f"Presentation saved: {os.path.join(main_folder, output_file)}")
