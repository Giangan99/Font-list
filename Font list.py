import os
import xlsxwriter
from PIL import Image, ImageDraw, ImageFont
import shutil

dir_path = r'E:\SETUP\Font AI+PTS\FONT\FontBase'

workbook = xlsxwriter.Workbook('D:/New folder/font_list.xlsx')
worksheet = workbook.add_worksheet('Fonts')
worksheet.write(0, 0, 'Font Name')
worksheet.write(0, 1, 'Font Image')
worksheet.write(0, 2, 'Font Path')
row = 1

# Create the fonts_images directory if it doesn't exist
fonts_images_dir = os.path.abspath('fonts_images')
if not os.path.exists(fonts_images_dir):
    os.makedirs(fonts_images_dir)

# Set up cell format for centering and border
cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})

for subdir, dirs, files in os.walk(dir_path):
    for file in files:
        if file.endswith('.ttf') or file.endswith('.otf'):
            font_path = os.path.join(subdir, file)
            font_name = os.path.splitext(file)[0]

            try:
                # Create the image file name and path
                image_name = font_name + ".jpg"
                image_path = os.path.join(fonts_images_dir, image_name)

                # Capture font image and save to disk
                font_sample = Image.new('RGB', (200, 80), color=(255, 255, 255))
                draw = ImageDraw.Draw(font_sample)
                font = ImageFont.truetype(font_path, 40)
                draw.text((10, 10), 'Font Sample', fill=(0, 0, 0), font=font)
                font_sample.save(image_path)

                # Insert font data into worksheet
                worksheet.write(row, 0, font_name)
                worksheet.insert_image(row, 1, image_path, {'x_offset': 15, 'y_offset': 10})
                worksheet.write(row, 2, font_path, cell_format)
                worksheet.set_column(2, 2, 15)
                worksheet.set_row(row, 80, None, {'border': 1})
                row += 1
            except OSError:
                print(f"Error: Unable to load font {font_path}")
                continue

# Set column widths and close workbook
worksheet.set_column('A:C', 25)
workbook.close()

# Delete temporary font images directory
shutil.rmtree(fonts_images_dir)
print("Done!")
