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

# Find all local image references
local_images = re.findall(r'src=["\']\.\/assets\/([^/"\']+)["\']', html_content)
local_images = set(local_images)  # Remove duplicates

print('Found local image references:')
print('=' * 80)
replacements = []
for img in sorted(local_images):
    # Extract clean name (filename without extension)
    clean_name = img.split('.')[0].lower()
    
    if clean_name in clean_name_map:
        cloudinary_url = clean_name_map[clean_name]
        print(f'✓ {img:20} -> {clean_name} -> Found in Excel')
        replacements.append((img, cloudinary_url, clean_name))
    else:
        print(f'✗ {img:20} -> {clean_name} -> NOT FOUND in Excel')

print(f'\nTotal local images: {len(local_images)}')
print(f'Successfully mapped: {len(replacements)}/{len(local_images)}')

# Now replace all occurrences
print('\n' + '=' * 80)
print('Performing replacements in HTML...')
print('=' * 80)

updated_html = html_content
for img, cloudinary_url, clean_name in replacements:
    # Replace both single and double quoted paths
    pattern1 = f'./assets/{img}'
    pattern2 = f"./assets/{img}"
    
    updated_html = updated_html.replace(f"src='{pattern1}'", f'src="{cloudinary_url}"')
    updated_html = updated_html.replace(f'src="{pattern1}"', f'src="{cloudinary_url}"')

# Write the updated HTML back to file
with open('index.html', 'w', encoding='utf-8') as f:
    f.write(updated_html)

print(f'\n✓ Successfully updated index.html with Cloudinary URLs!')
print(f'  - Total replacements made: {len(replacements)}')
print(f'  - File saved: index.html')
