from flask import Flask, request
from twilio.twiml.messaging_response import MessagingResponse
import pandas as pd
import xlwings as xw

app = Flask(__name__)

# Ruta al Excel
file_path = "BASE 202540 TAC-CHU (2).xlsm"

# Leer los datos una vez al arrancar
df = pd.read_excel(file_path, sheet_name="ID adeudos")
df.columns = df.columns.str.strip()

# Diccionario para controlar el estado de cada usuario (número de WhatsApp)
estados = {}

@app.route("/whatsapp", methods=["POST"])
def whatsapp():
    numero = request.form.get("From")  # Número del usuario
    mensaje = request.form.get("Body").strip()
    respuesta = MessagingResponse()
    msg = respuesta.message()

    # Obtener estado actual del usuario (si no hay, es nuevo)
    estado = estados.get(numero)

    # Si es nuevo o no tiene estado, saludar y pedir nombre
    if estado is None:
        if mensaje.lower() in ["hola", "hi", "buenos días", "buenas", "buenas tardes"]:
            estados[numero] = {"paso": 1}
            msg.body("👋 ¡Hola! Soy el asistente de UNID.\n\nPor favor, escribe tu *nombre completo* tal como aparece en el sistema para continuar.")
        else:
            estados[numero] = {"paso": 1, "nombre": mensaje.lower()}
            msg.body("✅ Gracias. Ahora ingresa tu *ID de alumno* para validar tus datos.")
        return str(respuesta)

    # Paso 1: ya se pidió el nombre, ahora espera el ID
    elif estado["paso"] == 1:
        estados[numero] = {"paso": 2, "nombre": mensaje.lower()}
        msg.body("✅ Gracias. Ahora ingresa tu *ID de alumno* para validar tus datos.")
        return str(respuesta)

    # Paso 2: validar nombre + ID
    elif estado["paso"] == 2:
        nombre_input = estado["nombre"]
        id_input = mensaje.strip()

        coincidencias = df[df['NOMBRE_LEGAL'].str.lower().str.strip() == nombre_input]

        if coincidencias.empty:
            estados.pop(numero, None)  # reiniciar
            msg.body("❌ El nombre no fue encontrado. Por favor, vuelve a escribirlo exactamente como aparece en el sistema.")
            return str(respuesta)

        alumno = coincidencias[coincidencias['ID_ALUMNO'].astype(str).str.strip() == id_input]

        if alumno.empty:
            estados.pop(numero, None)
            msg.body("❌ El ID no coincide con el nombre. Inicia de nuevo escribiendo tu *nombre completo*.")
            return str(respuesta)

        # Datos válidos
        row = alumno.iloc[0]
        nombre = row['NOMBRE_LEGAL']
        id_alumno = row['ID_ALUMNO']
        programa = row['PROGRAMA']
        campus = row['CAMPUS']
        adeudo = row['ADEUDO']

        # Guardar en Excel hoja AUX
        app_excel = xw.App(visible=False)
        wb = app_excel.books.open(file_path)
        hoja = wb.sheets["AUX"]

        hoja["A1"].value = "NOMBRE"
        hoja["B1"].value = nombre
        hoja["A2"].value = "ID"
        hoja["B2"].value = id_alumno
        hoja["A3"].value = "PROGRAMA"
        hoja["B3"].value = programa
        hoja["A4"].value = "CAMPUS"
        hoja["B4"].value = campus
        hoja["A5"].value = "ADEUDO"
        hoja["B5"].value = adeudo

        wb.save()
        wb.close()
        app_excel.quit()

        estados[numero] = {
            "paso": 3,
            "nombre": nombre,
            "id": id_alumno
        }

        msg.body(f"""🎓 *Datos verificados correctamente:*
👤 Nombre: {nombre}
🆔 ID: {id_alumno}
🏫 Campus: {campus}
📘 Programa: {programa}
💰 Adeudo: ${adeudo}

¿Deseas que genere tu ficha de pago en PDF? Responde *Sí* o *No*.
""")
        return str(respuesta)

    # Paso 3: generar o no el PDF
    elif estado["paso"] == 3:
        if mensaje.lower() in ["sí", "si"]:
            try:
                app_excel = xw.App(visible=False)
                wb = app_excel.books.open(file_path)
                wb.macro("GenerarFichaPDF")()
                wb.save()
                wb.close()
                app_excel.quit()
                msg.body("✅ Tu ficha fue generada correctamente. Pronto estará disponible para descargar.")
            except Exception as e:
                msg.body(f"⚠️ Error al generar el PDF: {e}")
        else:
            msg.body("👌 Entendido. No se generó la ficha.")

        estados.pop(numero, None)
        return str(respuesta)

    # Si el flujo se rompe
    estados.pop(numero, None)
    msg.body("⚠️ Ha ocurrido un error. Por favor escribe tu *nombre completo* para comenzar de nuevo.")
    return str(respuesta)

if __name__ == "__main__":
    app.run(debug=True)
