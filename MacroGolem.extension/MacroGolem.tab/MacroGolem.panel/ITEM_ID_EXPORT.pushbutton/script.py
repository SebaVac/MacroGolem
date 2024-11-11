# -*- coding: utf-8 -*-
__title__   = "ITEM_ID_EXPORT"
__doc__     = """Version = 1.0
Date    = 6.11.2024
________________________________________________________________
Description:

Download all element ids from a bridge and store them in an xlsx file.
________________________________________________________________
How-To:

1. Run the script.
2. The file will be saved automatically in the specified folder.
3. Check the generated file with all Element IDs.

________________________________________________________________
Author: Macro golem team"""

# Imports
from pyrevit import revit, DB
import xlsxwriter
from datetime import datetime
import os

# Main function
def export_element_ids():
    # Get the current Revit document
    doc = revit.doc

    # Get all elements in the document (including system families)
    collector = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()

    # Define the folder where the file will be saved in the user's Documents directory
    folder = os.path.join(os.path.expanduser("~"), "Documents", "Automatizacion_BIM_datos", "outputs")
    
    if not os.path.exists(folder):
        os.makedirs(folder)  # Create the folder if it doesn't exist

    # Define the Excel filename with a timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = os.path.join(folder, "revit_ids_{}.xlsx".format(timestamp))

    # Create a new Excel file and worksheet
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet("Element IDs")

    # Add headers
    worksheet.write(0, 0, "Element ID")
    worksheet.write(0, 1, "Category")
    worksheet.write(0, 2, "Family Name")
    worksheet.write(0, 3, "Type Name")

    # Populate data
    for idx, el in enumerate(collector, start=1):  # Start from row 1, since row 0 is for headers
        element_id = el.Id.IntegerValue
        category = el.Category.Name if el.Category else "No Category"
        family_name = el.LookupParameter("Family Name").AsString() if el.LookupParameter("Family Name") else "N/A"
        type_name = el.Name if hasattr(el, "Name") else "N/A"

        # Write data to the Excel sheet
        worksheet.write(idx, 0, element_id)
        worksheet.write(idx, 1, category)
        worksheet.write(idx, 2, family_name)
        worksheet.write(idx, 3, type_name)

    # Save the Excel file
    workbook.close()
    
    print("File saved at: {}".format(filename))

# Run the function
export_element_ids()
