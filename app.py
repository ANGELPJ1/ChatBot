# Librerias para:
# Flask y creacion de servidor web para puerto hacia TWILIO
# Twilio para conexion al servidor
from flask import Flask, request
from twilio.twiml.messaging_response import MessagingResponse

app = Flask(__name__)

# End Point POST para puebas iniciales
@app.route("/whatsapp", methods=["POST"])
def whatsapp_reply():
    incoming_msg = request.form.get('Body').strip()
    resp = MessagingResponse()
    msg = resp.message()

    # Ejemplo simple: responder según mensaje
    if "hola" in incoming_msg.lower():
        msg.body("¡Hola! Por favor, indícame tu nombre completo.")
    else:
        msg.body("Estoy en desarrollo, pero pronto podré ayudarte 😄")

    return str(resp)

if __name__ == "__main__":
    app.run(debug=True)
