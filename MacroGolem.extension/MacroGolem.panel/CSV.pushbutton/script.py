# -*- coding: utf-8 -*-
__title__   = "CSV"
__doc__     = """Version = 1.0
Date    = 15.06.2024
________________________________________________________________
Description:

This is the placeholder for a .pushbutton in a /pulldown
You can use it to start your pyRevit Add-In

________________________________________________________________
How-To:

1. [Hold ALT + CLICK] on the button to open its source folder.
You will be able to override this placeholder.

2. Automate Your Boring Work ;)

________________________________________________________________
TODO:
[FEATURE] - Describe Your ToDo Tasks Here
________________________________________________________________
Last Updates:
- [15.06.2024] v1.0 Change Description
- [10.06.2024] v0.5 Change Description
- [05.06.2024] v0.1 Change Description 
________________________________________________________________
Author: Erik Frits"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================
import clr
import csv
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, ElementId, StorageType

# Importar el módulo de Windows Forms
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog

# Obtener el documento activo de Revit
doc = __revit__.ActiveUIDocument.Document

# Función para permitir la selección de archivo mediante un cuadro de diálogo
def select_csv_file():
    dialog = OpenFileDialog()
    dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"  # Filtrar solo archivos .csv
    dialog.Title = "Seleccione el archivo CSV"
    dialog.ShowDialog()
    # Comprobar si se seleccionó un archivo
    if dialog.FileName:
        return dialog.FileName
    else:
        return None

# Función para actualizar los parámetros en Revit
def update_parameters_from_csv(doc):
    # Abrir el cuadro de diálogo para seleccionar el archivo
    csv_file_path = select_csv_file()
    if csv_file_path is None:
        print("No se seleccionó ningún archivo.")
        return

    # Abrir y leer el archivo CSV
    with open(csv_file_path, mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        
        # Iniciar transacción para modificar elementos en Revit
        with Transaction(doc, "Actualizar parámetros desde CSV") as transaction:
            transaction.Start()
            
            # Recorrer cada fila del CSV
            for row in csv_reader:
                element_id = int(row['ElementID'])  # Obtener ID del elemento desde el CSV
                element = doc.GetElement(ElementId(element_id))  # Buscar el elemento en el modelo Revit
                
                if element:
                    # Actualizar los parámetros del elemento desde el CSV
                    for param_name, value in row.items():
                        if param_name != 'ElementID':  # Saltar el ID del elemento
                            param = element.LookupParameter(param_name)
                            if param and param.StorageType == StorageType.String:
                                param.Set(value)
                            elif param and param.StorageType == StorageType.Double:
                                param.SetValueString(value)
                            elif param and param.StorageType == StorageType.Integer:
                                param.Set(int(value))
            
            # Confirmar cambios
            transaction.Commit()

# Llamar a la función para actualizar los parámetros
update_parameters_from_csv(doc)



#==================================================
#🚫 DELETE BELOW
#from Snippets._customprint import kit_button_clicked    # Import Reusable Function from 'lib/Snippets/_customprint.py'
#kit_button_clicked(btn_name=__title__)                  # Display Default Print Message
