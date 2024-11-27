# -*- coding: utf-8 -*-
__title__ = "Exportar Todo el Puente"
__doc__ = """Exporta todos los elementos del puente con traducción de valores específicos a un archivo CSV."""

import clr
import csv
import os
from Autodesk.Revit.DB import FilteredElementCollector, ElementId

# Obtener el documento activo en Revit
uiapp = __revit__  # pyRevit proporciona este objeto
doc = uiapp.ActiveUIDocument.Document

# Ruta de salida para el archivo CSV
output_path = "C:\\Users\\Usuario\\Documents\\Automatizacion_BIM_datos\\outputs\\Exportacion.csv"

# Lista de parámetros a exportar
parameters_to_export = [
    "Nivel de referencia", "Plano de trabajo", "Desfase de nivel inicial",
    "Desfase de nivel final", "Orientación", "Rotación de sección transversal",
    "Extensión inicial", "Extensión final", "Justificación YZ",
    "Justificación Y", "Valor de desfase Y", "Justificación Z", "Valor de desfase Z",
    "Material estructural", "Ubicación de símbolo de barras", "Conexión de inicio",
    "Conexión de fin", "Longitud de corte", "Uso estructural", "Longitud", "Volumen",
    "Elevación en parte superior", "Elevación en parte inferior", "Imagen", "Comentarios",
    "Marca", "Tiene asociación", "ID elemento", "Fase de creación", "Fase de derribo",
    "Tipo predefinido de IFC", "Exportar como IFC como", "Exportar a IFC", "IfcGUID", "NR"
]

# Mapeos de valores para parámetros específicos
value_maps = {
    "Justificación YZ": {0: "Uniforme", 1: "Independiente"},
    "Justificación Y": {0: "Origen", 1: "Parte Superior", 2: "Centro", 3: "Parte Inferior"},
    "Justificación Z": {0: "Origen", 1: "Parte Superior", 2: "Centro", 3: "Parte Inferior"},
    "Ubicación de símbolo de barras": {
        0: "Parte superior de geometría",
        1: "Centro de geometría",
        2: "Parte inferior de geometría",
        3: "Línea de ubicación"
    },
    "Conexión de inicio": {0: "Momento de estructura", 1: "Momento de voladizo", 2: "Ninguno"},
    "Conexión de fin": {0: "Momento de estructura", 1: "Momento de voladizo", 2: "Ninguno"},
    "Uso estructural": {
        0: "Jacena", 1: "Tornapunta horizontal", 2: "Viguetas", 3: "Otro", 4: "Correa"
    },
    "Fase de creación": {
        0: "Existente", 1: "Demostración", 2: "Fase 1", 3: "Fase 2", 4: "Fase 3"
    },
    "Fase de derribo": {
        0: "Existente", 1: "Demostración", 2: "Fase 1", 3: "Fase 2", 4: "Fase 3"
    }
}

# Crear los encabezados del CSV
headers = ["Element ID", "Category", "Family Name", "Type Name"] + parameters_to_export

# Función para exportar datos al archivo CSV
def export_bim_data(doc, output_path, headers):
    # Asegurar que el archivo no esté bloqueado
    if os.path.exists(output_path):
        try:
            os.remove(output_path)  # Eliminar el archivo si existe
        except Exception as e:
            raise IOError("El archivo está bloqueado o no se puede eliminar: {0}".format(output_path))

    # Colector para todos los elementos del modelo
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()

    # Abrir el archivo CSV para escritura
    with open(output_path, mode='w') as csvfile:
        csv_writer = csv.writer(csvfile, delimiter=';')

        # Escribir los encabezados
        csv_writer.writerow(headers)

        # Recorrer los elementos del modelo
        for element in collector:
            try:
                # Obtener información general del elemento
                element_id = element.Id.IntegerValue
                category = element.Category.Name if element.Category else "Sin Categoría"
                family_name = element.Symbol.FamilyName if hasattr(element, "Symbol") else "N/A"
                type_name = element.Name if hasattr(element, "Name") else "N/A"

                # Iniciar la fila con la información general
                row = [element_id, category, family_name, type_name]

                # Agregar valores de los parámetros especificados
                for param_name in parameters_to_export:
                    param = element.LookupParameter(param_name)
                    if param:
                        if param.StorageType.ToString() == "String":
                            row.append(param.AsString() or "Sin Valor")
                        elif param.StorageType.ToString() == "Double":
                            row.append(param.AsValueString() or "Sin Valor")
                        elif param.StorageType.ToString() == "Integer":
                            # Aplicar mapeo de valores si corresponde
                            if param_name in value_maps:
                                row.append(value_maps[param_name].get(param.AsInteger(), "Valor Desconocido"))
                            else:
                                row.append(param.AsInteger())
                        else:
                            row.append("No manejado")
                    else:
                        row.append("N/A")

                # Escribir la fila en el archivo CSV
                csv_writer.writerow(row)

            except Exception as e:
                # Manejo de errores
                print("Error procesando el elemento con ID {0}: {1}".format(element_id, e))

# Ejecutar la función de exportación
export_bim_data(doc, output_path, headers)

print("Datos exportados exitosamente a {0}".format(output_path))
