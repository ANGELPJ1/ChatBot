# Librerias:
# Pandas para la lectura y filtrado de datos
import pandas as pd
# xliwings para lectura de archivos excel con macros sin afectarlas
import xlwings as xw

# Archivo origen con los datos
file_path = "BASE 202540 TAC-CHU (2).xlsm"

# Lectura de hoja de adeudos usando pandas enfocandose en la hoja de ID ADEUDOS
df = pd.read_excel(file_path, sheet_name="ID adeudos")
df.columns = df.columns.str.strip()

while True:
    # Solicitud de nombre en el excel
    nombre_input = input("Ingrese el nombre completo del alumno: ").strip().lower()

    # Busqueda del nombre de la columna de NOMBRE_LEGAL
    coincidencias = df[df['NOMBRE_LEGAL'].str.lower().str.strip() == nombre_input]

    # Bifurcacion para casos posibles
    if coincidencias.empty:
        print("No se encontró ningún alumno con ese nombre. Inténtelo de nuevo.\n")
        continue

    # Busqueda de ID para validar alumno
    id_input = input("Ingrese el ID del alumno: ").strip()

    # Bifurcacion para casos posibles
    alumno = coincidencias[coincidencias['ID_ALUMNO'].astype(str).str.strip() == id_input]

    if alumno.empty:
        # Los datos no coinciden
        print("El ID no coincide con el nombre. Verifique los datos.\n")
        continue

    # Si los datos coinciden de agregan en:
    row = alumno.iloc[0]
    nombre = row['NOMBRE_LEGAL']
    id_alumno = row['ID_ALUMNO']
    programa = row['PROGRAMA']
    campus = row['CAMPUS']
    adeudo = row['ADEUDO']

    # Impresion de datos encontrados en sistema
    print("\n Datos encontrados:")
    print(f"Nombre: {nombre}")
    print(f"ID: {id_alumno}")
    print(f"Programa: {programa}")
    print(f"Campus: {campus}")
    print(f"Adeudo: {adeudo}")
    print("-" * 40)

    # Abrir el archivo con xlwings y pasar los datos a hoja AUX para generacion de PDF
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    sheet = wb.sheets["AUX"]

    # Escribir en hoja AUX para su plantilla y posterior paso a FICHA_PAGO con los datos encontrados
    sheet["A1"].value = "NOMBRE"
    sheet["B1"].value = nombre
    sheet["A2"].value = "ID"
    sheet["B2"].value = id_alumno
    sheet["A3"].value = "PROGRAMA"
    sheet["B3"].value = programa
    sheet["A4"].value = "CAMPUS"
    sheet["B4"].value = campus
    sheet["A5"].value = "ADEUDO"
    sheet["B5"].value = adeudo

    # Guardar y cerrar
    wb.save()
    wb.close()
    app.quit()

    # Impresion de prueba para verificar pasos anteriores
    print("📄 Datos escritos en la hoja AUX. Puedes generar la ficha desde Excel ahora.")

    # Pregunta al usuario si desea generar el PDF
    respuesta = input("¿Desea generar el PDF de la ficha de pago? (s/n): ").strip().lower()

    # En caso de que si reabrir Excel y ejecutar macro
    if respuesta == 's':
        app = xw.App(visible=False)
        wb = app.books.open(file_path)

        try:
            # Ejecucion de macro 'GenerarFichaPDF'
            wb.macro("GenerarFichaPDF")()
            print("PDF generado correctamente desde la macro.")
        except Exception as e:
            print("Error al ejecutar la macro:", e)

        wb.close()
        app.quit()
    else:
        print("Operación finalizada sin generar PDF.")

    break
