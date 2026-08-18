const ID_HOJA = "1u-vlocP2AIQI0tFKo0GFrSN5HH_20Qyar2rogG9oWZ4"; 
const URL_FORMULARIO_BASE = "https://docs.google.com/forms/d/e/1FAIpQLSf1jA5HPrlWWBTWkiTmGcKFU6LSCM4paDsqzAUJDiChCxH9tg/viewform?usp=pp_url&entry.1793374690=";


function doGet(e) {
  const accion = e ? e.parameter.accion : "";
  const codigo = e ? e.parameter.codigo : "";

  if (accion === "buscarLibro" && codigo) {
    const doc = SpreadsheetApp.openById(ID_HOJA);
    // Busca en la pestaña 'Inventario' o usa la primera pestaña disponible
    const sheetInventario = doc.getSheetByName("Inventario") || doc.getSheets()[0];
    const datos = sheetInventario.getDataRange().getValues();

    let tituloEncontrado = "";
    let encontrado = false;

    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim() === String(codigo).trim()) {
        tituloEncontrado = datos[i][1]; // Columna B (Título)
        encontrado = true;
        break;
      }
    }

    const resultado = {
      encontrado: encontrado,
      titulo: tituloEncontrado
    };

    return ContentService.createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput("Servicio de Biblioteca Activo");
}


function doPost(e) {
  try {
    const doc = SpreadsheetApp.openById(ID_HOJA);
    const sheetPrestamos = doc.getSheetByName("Hoja 1") || doc.getSheets()[0];

    const codigo = e.parameter.codigo;
    const libro = e.parameter.libro;
    const correo = e.parameter.correo;
    const fecha = e.parameter.fecha;

    let fechaFormateada = fecha;
    if (fecha && fecha.includes('-')) {
      const partes = fecha.split('-');
      fechaFormateada = `${partes[2]}/${partes[1]}/${partes[0]}`;
    }

    sheetPrestamos.appendRow([codigo, libro, correo, fechaFormateada, "Prestado", "No"]);

    return ContentService.createTextOutput("Éxito").setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput("Error: " + err.toString()).setMimeType(ContentService.MimeType.TEXT);
  }
}


function enviarRecordatorios2DiasAntes() {
  const doc = SpreadsheetApp.openById(ID_HOJA);
  const sheet = doc.getSheetByName("Hoja 1") || doc.getSheets()[0];
  const datos = sheet.getDataRange().getValues();
  
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  for (let i = 1; i < datos.length; i++) {
    const tituloLibro = datos[i][1];
    const emailUsuario = datos[i][2];
    const fechaRaw = datos[i][3];
    const estado = datos[i][4];
    const recordatorioEnviado = datos[i][5];

    if (estado === "Prestado" && recordatorioEnviado !== "Sí" && emailUsuario && fechaRaw) {
      let fechaDevolucion;
      if (fechaRaw instanceof Date) {
        fechaDevolucion = new Date(fechaRaw.getTime());
      } else if (typeof fechaRaw === 'string' && fechaRaw.includes('/')) {
        let partes = fechaRaw.split('/');
        fechaDevolucion = new Date(partes[2], partes[1] - 1, partes[0]);
      } else {
        fechaDevolucion = new Date(fechaRaw);
      }

      fechaDevolucion.setHours(0, 0, 0, 0);
      const diferencia = fechaDevolucion.getTime() - hoy.getTime();
      const diasFaltantes = Math.round(diferencia / (1000 * 60 * 60 * 24));

      if (diasFaltantes === 2) {
        const fechaTexto = `${String(fechaDevolucion.getDate()).padStart(2, '0')}/${String(fechaDevolucion.getMonth() + 1).padStart(2, '0')}/${fechaDevolucion.getFullYear()}`;
        const asunto = `⏰ Recordatorio: Devolución de "${tituloLibro}" en 2 días`;

        const urlAplazar = URL_FORMULARIO_BASE + (i + 1);
        
        const mensajePlano = `Hola,\n\nTe recordamos que faltan 2 días para devolver "${tituloLibro}".\nFecha límite: ${fechaTexto}\n\nSi deseas aplazar 3 días más haz clic aquí:\n${urlAplazar}`;

        const htmlBody = `
          <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 500px; padding: 24px; border: 1px solid #e2e8f0; border-radius: 12px; background-color: #ffffff; color: #1e293b;">
            <h2 style="color: #2563eb; margin-top: 0; font-size: 20px;">📚 Recordatorio de Biblioteca</h2>
            <p style="font-size: 15px; color: #334155;">Hola,</p>
            <p style="font-size: 15px; color: #334155;">Te recordamos que faltan <strong>2 días</strong> para la fecha límite de devolución del libro <strong>"${tituloLibro}"</strong>.</p>
            
            <div style="background-color: #f1f5f9; padding: 12px; border-radius: 8px; font-weight: bold; text-align: center; font-size: 16px; color: #0f172a; margin: 20px 0;">
              📅 Fecha límite actual: ${fechaTexto}
            </div>

            <p style="font-size: 14px; color: #64748b; text-align: center;">¿Necesitas más tiempo para terminar de leerlo?</p>
            
            <div style="text-align: center; margin: 20px 0;">
              <a href="${urlAplazar}" target="_blank" style="background-color: #2563eb; color: #ffffff; padding: 12px 24px; text-decoration: none; border-radius: 8px; font-weight: 600; font-size: 15px; display: inline-block;">
                ➕ Aplazar 3 días más
              </a>
            </div>

            <p style="color: #94a3b8; font-size: 12px; text-align: center; margin-top: 25px;">Si devuelves el libro a tiempo en la biblioteca, puedes ignorar este mensaje.</p>
          </div>
        `;

        GmailApp.sendEmail(emailUsuario, asunto, mensajePlano, { htmlBody: htmlBody });
        sheet.getRange(i + 1, 6).setValue("Sí");
      }
    }
  }
}


function procesarRespuestaFormulario(e) {
  try {
    var resp = e.values;
    var fila = parseInt(resp[1], 10); 

    if (!isNaN(fila) && fila > 1) {
      var doc = SpreadsheetApp.openById(ID_HOJA);
      var sheet = doc.getSheetByName("Hoja 1") || doc.getSheets()[0];

      var fechaActualVal = sheet.getRange(fila, 4).getValue();
      var fechaActual;

      if (fechaActualVal instanceof Date) {
        fechaActual = new Date(fechaActualVal.getTime());
      } else if (typeof fechaActualVal === 'string' && fechaActualVal.includes('/')) {
        var partes = fechaActualVal.split('/');
        fechaActual = new Date(partes[2], partes[1] - 1, partes[0]);
      } else {
        fechaActual = new Date(fechaActualVal);
      }

      fechaActual.setDate(fechaActual.getDate() + 3);

      var nuevaFechaTexto = String(fechaActual.getDate()).padStart(2, '0') + '/' + 
                            String(fechaActual.getMonth() + 1).padStart(2, '0') + '/' + 
                            fechaActual.getFullYear();

      sheet.getRange(fila, 4).setValue(nuevaFechaTexto);
      sheet.getRange(fila, 6).setValue("");
    }
  } catch (err) {
    Logger.log("Error al procesar formulario: " + err.toString());
  }
}
