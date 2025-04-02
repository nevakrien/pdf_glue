#!/home/user/anaconda3/envs/doc/bin/python
import os
import argparse
from os.path import join, isfile, isdir, basename
from PIL import Image
from docx import Document
from docx.shared import Inches

def process_images(directory):
    # Create a new Document
    doc = Document()
    
    # List only files in the directory
    images = [join(directory, x) for x in os.listdir(directory) if isfile(join(directory, x))]
    images.sort()
    
    # Maximum width in inches (page width - margins)
    max_width_inches = 6.0  # Adjust based on your document's margin settings

    # Process and add each image to the document
    for img in images:
        try:
            with Image.open(img) as im:
                width, height = im.size

                # Rotate the image if width is greater than height
                if width > height:
                    im = im.rotate(-90, expand=True)
                    width, height = im.size  # Update dimensions after rotation

                # Calculate width in inches using DPI if available (assumes 96 DPI if not)
                if 'dpi' in im.info:
                    dpi = im.info['dpi']
                    width_inches = width / dpi[0]
                else:
                    width_inches = width / 96

                # Scale image if its width exceeds the maximum width
                if width_inches > max_width_inches:
                    scale_factor = max_width_inches / width_inches
                    width_inches *= scale_factor
                    if 'dpi' in im.info:
                        dpi = im.info['dpi']
                        height_inches = (height / dpi[1]) * scale_factor
                    else:
                        height_inches = (height / 96) * scale_factor
                else:
                    if 'dpi' in im.info:
                        dpi = im.info['dpi']
                        height_inches = height / dpi[1]
                    else:
                        height_inches = height / 96

                # Save the (possibly rotated) image temporarily
                temp_path = f"{img}_rotated.jpg"
                im.save(temp_path)
            
            # Add the image to the document with calculated dimensions
            doc.add_picture(temp_path, width=Inches(width_inches), height=Inches(height_inches))
            
            # Remove the temporary image file
            os.remove(temp_path)
        except Exception as e:
            print(f"Error processing image {img}: {e}")
    
    # Save the document using the directory's basename
    output_filename = f'{basename(directory)}.docx'
    doc.save(output_filename)
    print(f"Document saved as {output_filename}")

def main():
    parser = argparse.ArgumentParser(
        description="Generate a Word document with images from a specified directory."
    )
    parser.add_argument("directory", help="Directory containing images to add to the document")
    args = parser.parse_args()
    
    if not os.path.exists(args.directory):
        print(f"Error: The directory '{args.directory}' does not exist.")
        exit(1)
    if not isdir(args.directory):
        print(f"Error: '{args.directory}' is not a directory.")
        exit(1)
    
    process_images(args.directory)

if __name__ == "__main__":
    main()
