function enviarCorreosProspectos() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var filas = hoja.getDataRange().getValues();

  for (var i = 1; i < filas.length; i++) {
    var fila = filas[i];

    var nombre = fila[0];   // A
    var email = fila[1];    // B
    var telefono = fila[3]; // D
    var estado = fila[6];   // G

    if (email == "" || estado == "Enviado") {
      continue;
    }

    try {
      // 1. Correo Interno (Simple)
      MailApp.sendEmail({
        to: "freddysilvatuesta@gmail.com",
        subject: "¡Nuevo Lead Registrado en MYPE X!",
        body: "Nuevo prospecto:\nNombre: " + nombre + "\nTeléfono: " + telefono
      });

      // 2. Correo al Prospecto (DISEÑO BONITO / TIPO TARJETA)
      var mensajeHtml = `
        <div style="background-color: #f4f4f4; padding: 40px 0; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;">
          <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; padding: 40px; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.1);">
            
            <h2 style="color: #333333; text-align: center; margin-top: 0;">¡Hola, ${nombre}!</h2>
            
            <p style="font-size: 16px; color: #555555; line-height: 1.6; text-align: center;">
              Gracias por tu interés en registrarte con nosotros. Hemos recibido tu información correctamente y estamos listos para ayudarte.
            </p>
            
            <p style="font-size: 16px; color: #555555; text-align: center; margin-bottom: 30px;">
              Elige cómo prefieres continuar:
            </p>

            <div style="text-align: center; margin-bottom: 20px;">
              <a href="https://www.fstnegocios.com" style="background-color: #007BFF; color: white; padding: 14px 25px; text-decoration: none; border-radius: 50px; font-weight: bold; font-size: 14px; margin: 5px; display: inline-block;">
                🌐 Visitar Web
              </a>

              <a href="https://wa.me/51949638568" style="background-color: #25D366; color: white; padding: 14px 25px; text-decoration: none; border-radius: 50px; font-weight: bold; font-size: 14px; margin: 5px; display: inline-block;">
                💬 Chatear por WhatsApp
              </a>
            </div>

            <hr style="border: 0; border-top: 1px solid #eeeeee; margin: 30px 0;">
            
            <p style="font-size: 12px; color: #999999; text-align: center;">
              Si tienes alguna duda, responde a este correo.<br>
              © 2025 MYPE X. Todos los derechos reservados.
            </p>
          </div>
        </div>
      `;

      MailApp.sendEmail({
        to: email,
        subject: "Confirmación de Registro - MYPE X",
        htmlBody: mensajeHtml
      });

      // 3. Escribir "Enviado" en Columna G (7)
      hoja.getRange(i + 1, 7).setValue("Enviado");

    } catch (e) {
      Logger.log("Error: " + e.toString());
    }
  }
}
