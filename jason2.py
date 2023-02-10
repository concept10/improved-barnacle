import pandas as pd
import xml.etree.ElementTree as ET

# Read the Excel spreadsheet into a pandas DataFrame
df = pd.read_excel("spreadsheet.xlsx")

# Extract the unique tag names and locations into separate lists
tag_names = df["Tag Name"].unique().tolist()
locations = df["Location"].tolist()

# Define a function to convert the location string into coordinates
def location_to_coordinates(location):
    # Extract the North and West zone information from the location string
    north_zone = int(location[:2])
    west_zone = location[2:]
    
    # Map the West zone letter to a column index
    west_zone_index = {"CA": 0, "BL": 1, "BK": 2, "BJ": 3}[west_zone]
    
    # Calculate the coordinates based on the North and West zones
    x = west_zone_index * 24
    y = (29 - north_zone) * 24
    
    return (x, y)

# Create the root element of the XML tree
root = ET.Element("BAX")

# Loop over the tag names and locations
for tag_name, location in zip(tag_names, locations):
    # Convert the location to coordinates
    x, y = location_to_coordinates(location)
    
    # Create a location element for the tag
    tag_location = ET.SubElement(root, "Location")
    tag_location.set("TagName", tag_name)
    tag_location.set("X", str(x))
    tag_


import os
import pandas as pd
import xml.etree.ElementTree as ET

# Read the Excel spreadsheet into a pandas DataFrame
df = pd.read_excel("spreadsheet.xlsx")

# Extract the unique tag names and locations into separate lists
tag_names = df["Tag Name"].unique().tolist()
locations = df["Location"].tolist()

# Define a function to convert the location string into coordinates
def location_to_coordinates(location):
    # Extract the North and West zone information from the location string
    north_zone = int(location[:2])
    west_zone = location[2:]
    
    # Map the West zone letter to a column index
    west_zone_index = {"CA": 0, "BL": 1, "BK": 2, "BJ": 3}[west_zone]
    
    # Calculate the coordinates based on the North and West zones
    x = west_zone_index * 24
    y = (29 - north_zone) * 24
    
    return (x, y)

# Loop over the unique tag names
for tag_name in tag_names:
    # Create a folder for the tag
    tag_folder = os.path.join("tags", tag_name)
    os.makedirs(tag_folder, exist_ok=True)
    
    # Create the root element of the XML tree
    root = ET.Element("BAX")
    
    # Loop over the tag names and locations
    for _tag_name, location in zip(df["Tag Name"], df["Location"]):
        if _tag_name == tag_name:
            # Convert the location to coordinates
            x, y = location_to_coordinates(location)
            
            # Create a location element for the tag
            tag_location = ET.SubElement(root, "Location")
            tag_location.set("TagName", tag_name)
            tag_location.set("X", str(x))
            tag_location.set("Y", str(y))
    
    # Write the XML tree to a file in the tag folder
    tree = ET.ElementTree(root)
    xml_path = os.path.join(tag_folder, "output.xml")
    tree.write(xml_path, xml_declaration=True, encoding="UTF-8", method="xml")

P