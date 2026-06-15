#!/usr/bin/env python
import openpyxl
import re

# Load the workbook
wb = openpyxl.load_workbook('Vishunu-mama_Cloudinary_Images.xlsx')
ws = wb.active

# Create mapping of clean names to Cloudinary URLs
clean_name_map = {}
for row in ws.iter_rows(min_row=2, max_row=100, values_only=True):
    if row[0] is None:
        break
    clean_name = str(row[2]).lower()  # Clean Name column, lowercase
    cloudinary_url = row[5]   # Cloudinary URL column
    clean_name_map[clean_name] = cloudinary_url

# Read the HTML file
with open('index.html', 'r', encoding='utf-8') as f:
    html_content = f.read()

# Find all CSS background image references
css_images = re.findall(r"url\(['\"]\./assets/([^/'\"]+)['\"]\)", html_content)
css_images = set(css_images)  # Remove duplicates

print('Found CSS background image references:')
print('=' * 80)
replacements = []
for img in sorted(css_images):
    # Extract clean name (filename without extension)
    clean_name = img.split('.')[0].lower()
    
    if clean_name in clean_name_map:
        cloudinary_url = clean_name_map[clean_name]
        print(f'✓ {img:20} -> {clean_name} -> Found in Excel')
        replacements.append((img, cloudinary_url, clean_name))
    else:
        print(f'✗ {img:20} -> {clean_name} -> NOT FOUND in Excel')

print(f'\nTotal CSS images: {len(css_images)}')
print(f'Successfully mapped: {len(replacements)}/{len(css_images)}')

# Now replace all CSS background image occurrences
print('\n' + '=' * 80)
print('Performing replacements in CSS background styles...')
print('=' * 80)

updated_html = html_content
for img, cloudinary_url, clean_name in replacements:
    # Replace in CSS background-image url() patterns
    pattern1 = f"url('./assets/{img}')"
    pattern2 = f'url("./assets/{img}")'
    
    updated_html = updated_html.replace(pattern1, f"url('{cloudinary_url}')")
    updated_html = updated_html.replace(pattern2, f'url("{cloudinary_url}")')

# Write the updated HTML back to file
with open('index.html', 'w', encoding='utf-8') as f:
    f.write(updated_html)

print(f'\n✓ Successfully updated CSS background image URLs!')
print(f'  - Total replacements made: {len(replacements)}')
print(f'  - File saved: index.html')
