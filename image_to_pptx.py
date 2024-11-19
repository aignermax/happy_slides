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

# Path to the main folder containing subdirectories with images
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

# Function to collect and sort images from a folder (including subfolders)
def get_images_sorted(folder):
    images = []
    for root, _, files in os.walk(folder):
        for file in files:
            if file.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif')):
                images.append(os.path.join(root, file))
    # Sort images in the current folder
    return natsorted(images)

# Sort main folders alphabetically
main_folders = [os.path.join(main_folder, d) for d in os.listdir(main_folder) if os.path.isdir(os.path.join(main_folder, d))]
main_folders = natsorted(main_folders)

# Process images from each main folder in the correct order
total_images = 0
for folder in main_folders:
    images = get_images_sorted(folder)
    total_images += len(images)

processed_images = 0
for folder in main_folders:
    print(f"Processing folder: {folder}")
    images = get_images_sorted(folder)
    for image_path in images:
        # Load image with correct orientation
        fixed_image = fix_image_orientation(image_path)
        if fixed_image is None:
            continue
        
        # Save image to an in-memory byte stream
        image_stream = io.BytesIO()
        fixed_image.save(image_stream, format='JPEG')
        image_stream.seek(0)  # Reset stream position to the beginning

        # Add a new slide for the image
        slide = presentation.slides.add_slide(presentation.slide_layouts[5])  # Blank slide
        slide_width = presentation.slide_width
        slide_height = presentation.slide_height

        # Add image and adjust size proportionally
        try:
            picture = slide.shapes.add_picture(image_stream, Inches(0), Inches(0), width=slide_width)
            picture.left = (slide_width - picture.width) // 2
            picture.top = (slide_height - picture.height) // 2
        except Exception as e:
            print(f"Error adding image: {image_path}. Error: {e}")
            continue

        processed_images += 1
        print(f"Processed {processed_images}/{total_images} images: {image_path}")

# Save the presentation
output_file = "Photos_Presentation.pptx"
presentation.save(output_file)
print(f"Presentation saved: {os.path.join(main_folder, output_file)}")
