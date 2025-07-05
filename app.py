# --------------------------------------------------------------
#                      ChatBot FLASK-WHATSAPP
# --------------------------------------------------------------
# Library needed:
# --------------------------------------------------------------
# Server flask to connect API
from flask import Flask, request
# Twilio to connect flask
from twilio.twiml.messaging_response import MessagingResponse
# Panda read an Excel file
import pandas as pd
# Xlwins read and execute an xlms without modifications
import xlwings as xw
# Clean text from Excel
import unicodedata
# --------------------------------------------------------------
# Load and import env
from dotenv import load_dotenv
import os
load_dotenv()
# Load var from use in Excel
EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH")
EXCEL_DATA_SHEET = os.getenv("EXCEL_DATA_SHEET")
EXCEL_AUX_SHEET = os.getenv("EXCEL_AUX_SHEET")
EXCEL_MACRO_NAME = os.getenv("EXCEL_MACRO_NAME")
# Load cols to use from Excel
COL_ID = os.getenv("COL_ID_ALUMNO")
COL_NOMBRE = os.getenv("COL_NOMBRE_LEGAL")
COL_PROGRAMA = os.getenv("COL_PROGRAMA")
COL_CAMPUS = os.getenv("COL_CAMPUS")
COL_ADEUDO = os.getenv("COL_ADEUDO")

app = Flask(__name__)

# Read DB
df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=EXCEL_DATA_SHEET)
df.columns = df.columns.str.strip()

# Normalize text
def limpiar(texto):
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    texto = " ".join(texto.split())
    return texto

# State control by number phone
estados = {}

@app.route("/whatsapp", methods=["POST"])
def whatsapp():
    numero = request.form.get("From")
    mensaje = request.form.get("Body").strip()
    respuesta = MessagingResponse()
    msg = respuesta.message()
    estado = estados.get(numero)
    mensaje_limpio = limpiar(mensaje)

    # Step 1
    if estado is None:
        if mensaje_limpio in ["hola", "hi", "buenos dias", "buenas tardes", "buenas noches", "buenas"]:
            estados[numero] = {"paso": 1}
            msg.body("👋 ¡Hola! Soy el asistente de UNID.\n\nPor favor, escribe tu *NOMBRE COMPLETO* tal como aparece en el sistema.")
        else:
            estados[numero] = {"paso": 2, "nombre": mensaje_limpio}
            msg.body("✅ Gracias. Ahora escribe tu *ID de alumno* para validar tus datos.")
        return str(respuesta)

    # Step 1: Get name
    if estado["paso"] == 1:
        estados[numero]["nombre"] = mensaje_limpio
        estados[numero]["paso"] = 2
        msg.body("✅ Gracias. Ahora escribe tu *ID de alumno* para validar tus datos.")
        return str(respuesta)

    # Step 2: Validate name + ID
    elif estado["paso"] == 2:
        nombre = estados[numero]["nombre"]
        id_input = mensaje.strip()

        coincidencias = df[df[COL_NOMBRE].apply(lambda x: limpiar(x)) == nombre]

        if coincidencias.empty:
            estados.pop(numero)
            msg.body("❌ No encontré ese nombre. Asegúrate de escribirlo como aparece en el sistema.\n\nEscribe *Hola* para intentarlo de nuevo.")
            return str(respuesta)

        alumno = coincidencias[coincidencias[COL_ID].astype(str).str.strip() == id_input]

        if alumno.empty:
            estados.pop(numero)
            msg.body("❌ El ID no coincide con el nombre, operacion fallida.\n\n Escribe *Hola* para intentar de nuevo.")
            return str(respuesta)

        row = alumno.iloc[0]
        estados[numero].update({
            "paso": 3,
            "id": id_input,
            "nombre_real": row[COL_NOMBRE],
            "programa": row[COL_PROGRAMA],
            "campus": row[COL_CAMPUS],
            "adeudo": row[COL_ADEUDO]
        })

        # Write in sheet AUX
        try:
            app_excel = xw.App(visible=False)
            wb = app_excel.books.open(EXCEL_FILE_PATH)
            hoja = wb.sheets[EXCEL_AUX_SHEET]
            hoja["A1"].value = "NOMBRE"
            hoja["B1"].value = row[COL_NOMBRE]
            hoja["A2"].value = "ID"
            hoja["B2"].value = id_input
            hoja["A3"].value = "PROGRAMA"
            hoja["B3"].value = row[COL_PROGRAMA]
            hoja["A4"].value = "CAMPUS"
            hoja["B4"].value = row[COL_CAMPUS]
            hoja["A5"].value = "ADEUDO"
            hoja["B5"].value = row[COL_ADEUDO]
            wb.save()
            wb.close()
            app_excel.quit()
        except Exception as e:
            msg.body(f"⚠️ Error al escribir en Excel: {e}")
            estados.pop(numero)
            return str(respuesta)

        msg.body(f"""🎓 *Datos encontrados:*
👤 Nombre: {row[COL_NOMBRE]}
🆔 ID: {id_input}
🏫 Campus: {row[COL_CAMPUS]}
📘 Programa: {row[COL_PROGRAMA]}
💰 Adeudo: ${row[COL_ADEUDO]}

¿Deseas generar tu ficha de pago en PDF? Responde *Sí* o *No*.
""")
        return str(respuesta)

    # Step 3: Generate PDF
    elif estado["paso"] == 3:
        if mensaje_limpio in ["si", "sí"]:
            msg.body("🛠️ Generando tu ficha, por favor espera...")

            try:
                app_excel = xw.App(visible=False)
                wb = app_excel.books.open(EXCEL_FILE_PATH)
                wb.macro(EXCEL_MACRO_NAME)()
                wb.save()
                wb.close()
                app_excel.quit()
                respuesta_final = "✅ Tu ficha fue generada correctamente."
            except Exception as e:
                respuesta_final = f"⚠️ Ocurrió un error al generar el PDF: {e}"
        else:
            respuesta_final = "👌 Entendido. No se generó la ficha."

        estados[numero]["paso"] = 4
        msg.body(respuesta_final + "\n\n🧾 ¿Deseas volver a ver tu información? Responde *Sí* o *No*.")
        return str(respuesta)

    # Step 4: Repeat or close
    elif estado["paso"] == 4:
        if mensaje_limpio in ["si", "sí"]:
            datos = estados[numero]
            msg.body(f"""🎓 *Datos encontrados:*
    👤 Nombre: {datos["nombre_real"]}
    🆔 ID: {datos["id"]}
    🏫 Campus: {datos["campus"]}
    📘 Programa: {datos["programa"]}
    💰 Adeudo: ${datos["adeudo"]}

    ✅ Gracias por consultar tu información. Si necesitas más ayuda, escribe *Hola*.""")
            estados.pop(numero)
        else:
            estados.pop(numero)
            msg.body("✅ Gracias por usar el asistente UNID. ¡Hasta pronto!")
        return str(respuesta)

if __name__ == "__main__":
    app.run(debug=True)
