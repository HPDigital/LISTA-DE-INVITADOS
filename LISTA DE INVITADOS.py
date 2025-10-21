"""
LISTA DE INVITADOS
"""

#!/usr/bin/env python
# coding: utf-8

# In[6]:


import pandas as pd

def load_guest_list_from_sheet(file_path, sheet_name):
    """
    Carga la lista de invitados desde una hoja específica.
    """
    sheet_data = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    # Los invitados se asumen desde la celda B3
    invitados = sheet_data.iloc[2:, 1:3]
    invitados.columns = ["Invitado1", "Invitado2"]
    # Añade el nombre de la hoja como una columna
    invitados.insert(0, "Nombre_Hoja", sheet_name)
    return invitados

def consolidate_guest_lists(file_path, excluded_sheet="LISTA UNIFICADA"):
    """
    Consolida las listas de invitados de todas las hojas excepto la excluida.
    """
    # Crea un DataFrame vacío para la lista consolidada
    consolidated_df = pd.DataFrame(columns=["Nombre_Hoja", "Invitado1", "Invitado2"])

    # Lee el archivo Excel y obtiene los nombres de las hojas
    excel_data = pd.ExcelFile(file_path)

    for sheet_name in excel_data.sheet_names:
        if sheet_name != excluded_sheet:
            invitados = load_guest_list_from_sheet(file_path, sheet_name)
            consolidated_df = pd.concat([consolidated_df, invitados], ignore_index=True)

    return consolidated_df

# Ruta al archivo Excel
file_path = r"C:\Users\HP\Downloads\Invitados.xlsx"

# Consolidar las listas de todas las hojas
lista_unificada_df_all = consolidate_guest_lists(file_path)

# Mostrar la lista consolidada al usuario
# Reemplaza esta línea por una opción que funcione en tu entorno
print("LISTA UNIFICADA con Todas las Hojas:")
print(lista_unificada_df_all)

# O guardar el DataFrame en un archivo Excel o CSV
output_path = r"C:\Users\HP\Downloads\LISTA_UNIFICADA.xlsx"
lista_unificada_df_all.to_excel(output_path, index=False)
print(f"Archivo guardado en: {output_path}")


# In[ ]:






if __name__ == "__main__":
    pass
