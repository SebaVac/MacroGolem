# -*- coding: utf-8 -*-
__title__   = "VR"
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
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, ElementId
import socket

# Configuración de socket para recibir datos de VR
HOST = '127.0.0.1'  # Dirección del servidor VR (localhost)
PORT = 65432        # Puerto para la comunicación

# Conectar con el socket del entorno VR
def receive_vr_data():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.connect((HOST, PORT))
        while True:
            data = s.recv(1024)
            if not data:
                break
            return data.decode('utf-8')  # Retorna los datos como string

# Actualizar parámetros de Revit con datos desde VR
def update_parameters_from_vr(doc):
    vr_data = receive_vr_data()  # Recibe los datos desde el entorno VR (simulados como string)
    
    # Simulación de cómo se recibirían los datos del VR: ElementID y parámetros
    vr_data = vr_data.split(',')
    element_id = int(vr_data[0])  # Suponemos que el primer dato es el ElementID
    param_name = vr_data[1]  # Nombre del parámetro
    param_value = vr_data[2]  # Valor del parámetro
    
    # Iniciar transacción para modificar el modelo en Revit
    with Transaction(doc, "Actualizar parámetros desde VR") as transaction:
        transaction.Start()
        
        # Buscar el elemento en Revit usando el ID
        element = doc.GetElement(ElementId(element_id))
        if element:
            param = element.LookupParameter(param_name)
            if param and param.StorageType == StorageType.String:
                param.Set(param_value)
            elif param and param.StorageType == StorageType.Double:
                param.SetValueString(param_value)
            elif param and param.StorageType == StorageType.Integer:
                param.Set(int(param_value))
        
        # Confirmar cambios
        transaction.Commit()

# Obtener el documento activo de Revit
doc = __revit__.ActiveUIDocument.Document

# Llamar a la función para recibir datos del entorno VR y actualizar el modelo en Revit
update_parameters_from_vr(doc)

#==================================================
#🚫 DELETE BELOW
from Snippets._customprint import kit_button_clicked    # Import Reusable Function from 'lib/Snippets/_customprint.py'
kit_button_clicked(btn_name=__title__)                  # Display Default Print Message
