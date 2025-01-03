Have you ever wanted to create a ppt for your Kindergarden with hundreds of photos? Well now you can. Simply use happy_slides!

# happy_slides

A Python script that scans directories for images, corrects their orientation based on EXIF data, and compiles them into a PowerPoint presentation with automatic slide transitions.

![HappySlides Logo](https://github.com/user-attachments/assets/3a480b87-5656-407c-8bfb-fae1b6fa5503)

## Features

- **Image Processing:** Automatically corrects image orientation using EXIF metadata.
- **Presentation Creation:** Inserts each image into individual slides in a PowerPoint presentation.
- **Sorting:** Organizes images based on folder structure and filenames.
- **Transitions:** Adds automatic slide transitions with a customizable duration.

## Requirements

- Python 3.x
- Required Python packages:
  - `python-pptx`
  - `Pillow`
  - `natsort`
  - `lxml`

## Installation

1. Clone or download this repository.
2. Install the required Python packages using pip:

   ```bash
   pip install python-pptx Pillow natsort lxml
   ```
## Usage
Place the script in the main directory containing the images and subdirectories you want to process.

## Run the script:

```bash
python happy_slides.py
```

`The script will generate a PowerPoint file named Photos_Presentation.pptx in the same directory.`

## Notes
Ensure that the images have correct EXIF orientation data for accurate processing. (usually the case)
The script processes common image formats such as .png, .jpg, .jpeg, .bmp, and .gif.
