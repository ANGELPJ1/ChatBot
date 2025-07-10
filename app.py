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
# Clean text from Excel
import unicodedata
# Move PDF to static
import shutil
# generate PDF
from fpdf import FPDF
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
RAILWAY_DOMAIN = os.getenv("RAILWAY_ENV", "http://localhost:5000")

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

    # Step 0: Initialize the process and attempts
    if estado is None:
        if mensaje_limpio in ["hola", "hi", "buenos dias", "buenas tardes", "buenas noches", "buenas"]:
            estados[numero] = {"paso": 1, "intentos": 0}
            msg.body(
                "👋 ¡Hola! Soy el asistente de UNID.\n\nPor favor, escribe tu *NOMBRE COMPLETO* tal como aparece en el sistema.")
        else:
            estados[numero] = {"paso": 2, "nombre": mensaje_limpio, "intentos": 0}
            msg.body("✅ Gracias. Ahora escribe tu *ID de alumno* para validar tus datos.")
        return str(respuesta)

    # Step 1: Get name
    if estado["paso"] == 1:
        nombre = mensaje_limpio
        coincidencias = df[df[COL_NOMBRE].apply(lambda x: limpiar(x)) == nombre]

        if coincidencias.empty:
            estados[numero]["intentos"] += 1
            if estados[numero]["intentos"] >= 3:
                msg.body("❌ Has superado el número máximo de intentos. Escribe *Hola* para comenzar de nuevo.")
                estados.pop(numero)
            else:
                msg.body(
                    f"❌ No encontré ese nombre. ({estados[numero]['intentos']}/3 intentos)\n\nPor favor, vuelve a escribir tu *NOMBRE COMPLETO* tal como aparece en el sistema.")
            return str(respuesta)

        # If there is coincidence
        estados[numero]["nombre"] = nombre
        estados[numero]["paso"] = 2
        estados[numero]["intentos"] = 0  # Initialize
        msg.body("✅ Gracias. Ahora escribe tu *ID de alumno* para validar tus datos.")
        return str(respuesta)


    # Step 2: Validate name + ID
    elif estado["paso"] == 2:
        nombre = estados[numero]["nombre"]
        id_input = mensaje.strip()

        coincidencias = df[df[COL_NOMBRE].apply(lambda x: limpiar(x)) == nombre]

        if coincidencias.empty:
            msg.body("❌ No encontré ese nombre. Escribe *Hola* para comenzar de nuevo.")
            estados.pop(numero)
            return str(respuesta)

        alumno = coincidencias[coincidencias[COL_ID].astype(str).str.strip() == id_input]

        if alumno.empty:
            estados[numero]["intentos"] += 1
            if estados[numero]["intentos"] >= 3:
                msg.body("❌ Has superado el número máximo de intentos. Escribe *Hola* para comenzar de nuevo.")
                estados.pop(numero)
            else:
                msg.body(
                    f"❌ El ID no coincide con el nombre. ({estados[numero]['intentos']}/3 intentos)\n\nPor favor, vuelve a escribir tu *ID de alumno* para validar los datos.")
            return str(respuesta)

        # If there is coincidence
        row = alumno.iloc[0]
        estados[numero].update({
            "paso": 3,
            "id": id_input,
            "nombre_real": row[COL_NOMBRE],
            "programa": row[COL_PROGRAMA],
            "campus": row[COL_CAMPUS],
            "adeudo": row[COL_ADEUDO],
            "intentos": 0  # Initialize
        })

        msg.body(f"""🎓 *Datos encontrados:*
    👤 Nombre: {row[COL_NOMBRE]}
    🆔 ID: {id_input}
    🏫 Campus: {row[COL_CAMPUS]}
    📘 Programa: {row[COL_PROGRAMA]}
    💰 Adeudo total: ${row[COL_ADEUDO]}

    ¿Deseas generar tu ficha de pago en PDF? Responde *Sí* o *No*.
    """)
        return str(respuesta)

    # Step 3: Generate PDF
    elif estado["paso"] == 3:
        if mensaje_limpio in ["si", "sí"]:
            msg.body("🛠️ Generando tu ficha, por favor espera...")

            # Generate PDF with fpdf
            try:
                id_alumno = estados[numero]["id"]
                nombre_pdf = f"Ficha_{id_alumno}.pdf"
                ruta_pdf = os.path.join("static", nombre_pdf)

                # Crear carpeta static si no existe
                os.makedirs("static", exist_ok=True)

                # Generar PDF con FPDF
                pdf = FPDF()
                pdf.add_page()
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.set_font("Arial", size=12)

                pdf.cell(200, 10, txt="Ficha de Pago UNID", ln=True, align="C")
                pdf.ln(10)

                datos = estados[numero]
                pdf.cell(200, 10, txt=f"Nombre: {datos['nombre_real']}", ln=True)
                pdf.cell(200, 10, txt=f"ID: {id_alumno}", ln=True)
                pdf.cell(200, 10, txt=f"Programa: {datos['programa']}", ln=True)
                pdf.cell(200, 10, txt=f"Campus: {datos['campus']}", ln=True)
                pdf.cell(200, 10, txt=f"Adeudo: ${datos['adeudo']}", ln=True)

                pdf.output(ruta_pdf)

                estados[numero]["paso"] = 4
                estados[numero]["pdf"] = nombre_pdf

                msg.body("✅ Ficha generada.\n\n¿Dónde deseas recibirla?\n\n👉 *WhatsApp* o *Correo*")
                return str(respuesta)

            except Exception as e:
                msg.body(f"⚠️ Error al generar el PDF: {e}")
                estados.pop(numero)
                return str(respuesta)
        else:
            estados[numero]["paso"] = 5
            msg.body("👌 La ficha no se generó.\n\n¿Deseas ver la información de nuevo? *Sí* o *No*.")
            return str(respuesta)

    # Step 4: Repeat or close
    elif estado["paso"] == 4:
        if mensaje_limpio == "whatsapp":
            nombre_pdf = estados[numero].get("pdf")
            if nombre_pdf:
                url_pdf = f"{RAILWAY_DOMAIN}/static/{nombre_pdf}"
                msg.body("📎 Aquí tienes tu ficha de pago:\n" + url_pdf)
                respuesta.message("✅ Gracias por usar el asistente UNID. ¡Hasta pronto!")

            else:
                msg.body("⚠️ No se encontró el archivo PDF. Intenta generar de nuevo con *Hola*.")

            # Reset the user state
            estados.pop(numero)
            return str(respuesta)

        elif mensaje_limpio in ["correo", "email"]:
            msg.body("📧 Tu ficha ha sido enviada por correo. Por favor, revisa tu bandeja de entrada.")

            # Closing and restarting message
            respuesta.message("✅ Gracias por usar el asistente UNID. ¡Hasta pronto!")
            estados.pop(numero)
            return str(respuesta)

        else:
            msg.body("❓ Responde *WhatsApp* o *Correo*.")
            return str(respuesta)

    # Step 5: Solicit information again or finish the process
    elif estado["paso"] == 5:
        # In case the information is requested again
        if mensaje_limpio in ["si", "sí"]:
            row = df[df[COL_ID].astype(str).str.strip() == estados[numero]["id"]].iloc[0]
            msg.body(f"""🎓 *Datos del alumno:*
    👤 Nombre: {row[COL_NOMBRE]}
    🆔 ID: {estados[numero]["id"]}
    🏫 Campus: {row[COL_CAMPUS]}
    📘 Programa: {row[COL_PROGRAMA]}
    💰 Adeudo: ${row[COL_ADEUDO]}

    📝 Si deseas volver a generar la ficha o consultar otra, escribe *Hola* nuevamente.""")
            estados.pop(numero)
        else:
            estados.pop(numero)
            msg.body("✅ Gracias por usar el asistente UNID. ¡Hasta pronto!")
        return str(respuesta)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))