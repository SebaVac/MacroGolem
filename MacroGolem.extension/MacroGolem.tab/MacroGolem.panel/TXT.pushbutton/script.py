# -*- coding: utf-8 -*-
__title__   = "TXT"
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

import clr
import csv
from Autodesk.Revit.DB import Transaction, ElementId, StorageType

# Importar el módulo de Windows Forms
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog

# Obtener el documento activo de Revit
doc = __revit__.ActiveUIDocument.Document

# Función para permitir la selección de archivo mediante un cuadro de diálogo
def select_txt_file():
    dialog = OpenFileDialog()
    dialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"  # Filtrar solo archivos .txt
    dialog.Title = "Seleccione el archivo TXT"
    dialog.ShowDialog()
    # Comprobar si se seleccionó un archivo
    if dialog.FileName:
        return dialog.FileName
    else:
        return None

# Función para leer el archivo .txt y extraer datos
def parse_txt_file(txt_file_path):
    elements_data = {}
    current_element = None
    
    with open(txt_file_path, 'r') as file:
        for line in file:
            line = line.strip()
            
            # Ignorar líneas vacías
            if not line:
                continue
            
            # Detectar nueva sección de elemento
            if "R:" in line:
                current_element = line.split(":")[0].strip()
                elements_data[current_element] = {}
            elif "ResultadoVF:" in line:
                # Asignar ResultadoVF al elemento actual
                elements_data[current_element]["ResultadoVF"] = line.split(":", 1)[1].strip()
            else:
                # Parsear otros datos de parámetros, dividiendo solo en el primer ":"
                parts = line.split(":", 1)
                if len(parts) == 2:
                    param_name, param_value = parts[0].strip(), parts[1].strip()
                    elements_data[current_element][param_name] = param_value
    
    return elements_data


# Función para actualizar los parámetros en Revit
def update_parameters_from_txt(doc, elements_data):
    with Transaction(doc, "Actualizar parámetros desde TXT") as transaction:
        transaction.Start()
        
        for element_name, params in elements_data.items():
            # Aquí deberás adaptar el proceso para encontrar el elemento en Revit según su nombre
            # Ejemplo: Buscar por nombre de familia, ID, o algún otro parámetro único en tu modelo
            element = None  # Deberás personalizar esto para obtener el elemento en Revit
            
            if element:
                # Actualizar los parámetros en el modelo
                for param_name, value in params.items():
                    param = element.LookupParameter(param_name)
                    if param:
                        if param.StorageType == StorageType.String:
                            param.Set(value)
                        elif param.StorageType == StorageType.Double:
                            param.SetValueString(value)
                        elif param.StorageType == StorageType.Integer:
                            param.Set(int(value))
        
        # Confirmar los cambios
        transaction.Commit()

# Seleccionar el archivo TXT
txt_file_path = select_txt_file()

# Verificar si el usuario seleccionó un archivo
if txt_file_path:
    # Leer los datos del archivo TXT
    elements_data = parse_txt_file(txt_file_path)

    # Actualizar los parámetros en Revit con los datos extraídos
    update_parameters_from_txt(doc, elements_data)
else:
    print("No se seleccionó ningún archivo.")



#==================================================
#🚫 DELETE BELOW
#from Snippets._customprint import kit_button_clicked    # Import Reusable Function from 'lib/Snippets/_customprint.py'
#kit_button_clicked(btn_name=__title__)                  # Display Default Print Message
