import pandas as pd

nombre_tabla_vieja = 'CONTACTOS_VIEJOS.xlsx'
nombre_tabla_nueva = 'DANEC.xlsx'
nombre_campo_viejo = 'NIF'
nombre_campo_nuevo = 'CEDULA_O_RUC'

try:
    archivo_excel_viejos = pd.read_excel(
        nombre_tabla_vieja, sheet_name='viejos')
    archivo_excel_nuevos = pd.read_excel(
        nombre_tabla_nueva, sheet_name='nuevos')
    print('----------------------FILES UPLOADED-----------------------------')
    print('----------------------PROCCESS STARED-----------------------------')
    # NUEVOS NO ESTAN DUPLICADOS
    nuevos_no_viejos = archivo_excel_nuevos[~archivo_excel_nuevos[nombre_campo_nuevo].isin(
        archivo_excel_viejos[nombre_campo_viejo])]
    # DATOS QUE SE DUPLICAN
    nuevos_duplicados = archivo_excel_nuevos[archivo_excel_nuevos[nombre_campo_nuevo].isin(
        archivo_excel_viejos[nombre_campo_viejo])]

    # DATOS DUPLICADOS EN SI MISMO
    nuevos_duplicados_en_ella = archivo_excel_nuevos[archivo_excel_nuevos[nombre_campo_nuevo].duplicated(keep=False)]
    no_duplicados_en_tabla = archivo_excel_nuevos[~archivo_excel_nuevos.duplicated(subset=[nombre_campo_nuevo], keep=False)]

    # GUARDANDO
    with pd.ExcelWriter('resultados.xlsx') as writer:
        nuevos_no_viejos.to_excel(writer, sheet_name='no_duplicados', index=False)
        nuevos_duplicados.to_excel(writer, sheet_name='duplicados', index=False)
        no_duplicados_en_tabla.to_excel(writer, sheet_name='omitidos_duplicados_ella_misma', index=False)
        nuevos_duplicados_en_ella.to_excel(writer, sheet_name='duplicados_ella_misma', index=False)

    print('----------------------FINISHED SAVED-----------------------------')

except Exception as error:
    print('----------------------ERROR-----------------------------')
    print(error)
