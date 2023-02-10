import xlrd
import os
import xml.etree.ElementTree as ET
import pytesseract
from PIL import Image

# Open the "Tracker" spreadsheet
workbook = xlrd.open_workbook('Tracker.xlsx')
worksheet = workbook.sheet_by_index(0)

# Create a data structure for each tagname
tag_data = []
for i in range(worksheet.nrows):
    tagname = worksheet.cell(i, 0).value
    # Generate a unique ID for each tagname
    tag_id = hash(tagname)
    # Generate a RGBA color for each tagname
    color = (i*20 % 256, i*30 % 256, i*40 % 256, 1)
    tag_data.append({'tagname': tagname, 'id': tag_id, 'color': color})

# Open the "pmd" spreadsheet
pmd_workbook = xlrd.open_workbook('pmd.xlsx')
pmd_worksheet = pmd_workbook.sheet_by_index(0)

# Associate each tagname in the "pmd" spreadsheet with the tagnames in the "Tracker" spreadsheet
for i in range(pmd_worksheet.nrows):
    pmd_tagname = pmd_worksheet.cell(i, 0).value
    for tag_info in tag_data:
        if pmd_tagname == tag_info['tagname']:
            tag_info['pdf_file'] = pmd_worksheet.cell(i, 1).value

# Search the workDirectory for each PDF file associated with each tagname
work_directory = '/path/to/workDirectory'
for tag_info in tag_data:
    pdf_file = tag_info['pdf_file']
    pdf_file_path = os.path.join(work_directory, pdf_file)
    # Use OCR to find the location in the PDF file
    img = Image.open(pdf_file_path)
    text = pytesseract.image_to_string(img)
    tag_info['location'] = text

# Generate the XML output
root = ET.Element("root")
for tag_info in tag_data:
    tag = ET.SubElement(root, "tag")
    ET.SubElement(tag, "tagname").text = tag_info['tagname']
    ET.SubElement(tag, "id").text = str(tag_info['id'])
    color = tag_info['color']
    ET.SubElement(tag, "color").text = "rgba({}, {}, {}, {})".format(color[0], color[1], color[2], color[3])
    ET.SubElement(tag, "location").text = tag_info['location']

tree = ET.ElementTree(root)
tree.write("output.xml")
