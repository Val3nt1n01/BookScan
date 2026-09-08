const ID_HOJA = "1u-vlocP2AIQI0tFKo0GFrSN5HH_20Qyar2rogG9oWZ4"; 

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'aplazar' && e.parameter.fila) {
    const fila = parseInt(e.parameter.fila, 10);
    const resultado = aplazarPrestamoDirecto(fila);
    
    const html = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <style>
          body { font-family: 'Segoe UI', Arial, sans-serif; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; background-color: #f1f5f9; }
          .card { background: white; padding: 32px; border-radius: 16px; box-shadow: 0 10px 25px rgba(0,0,0,0.1); text-align: center; max-width: 400px; }
          h2 { color: #0f172a; margin-top: 0; }
          p { color: #334155; font-size: 16px; margin: 20px 0; }
        </style>
      </head>
      <body>
        <div class="card">
          <h2>📚 Gestión de Biblioteca</h2>
          <p>${resultado}</p>
        </div>
      </body>
      </html>
    `;
    return HtmlService.createHtmlOutput(html);
  }

  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Registro de Préstamos - Biblioteca')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function aplazarPrestamoDirecto(fila) {
  try {
    if (isNaN(fila) || fila <= 1) return "Fila no válida.";
    
    const doc = SpreadsheetApp.openById(ID_HOJA);
    const sheet = doc.getSheetByName("Hoja 1") || doc.getSheets()[0];

    const fechaActualVal = sheet.getRange(fila, 4).getValue();
    let fechaActual;

    if (fechaActualVal instanceof Date) {
      fechaActual = new Date(fechaActualVal.getTime());
    } else if (typeof fechaActualVal === 'string' && fechaActualVal.includes('/')) {
      const partes = fechaActualVal.split('/');
      fechaActual = new Date(partes[2], partes[1] - 1, partes[0]);
    } else {
      fechaActual = new Date(fechaActualVal);
    }

    fechaActual.setDate(fechaActual.getDate() + 3);

    const nuevaFechaTexto = String(fechaActual.getDate()).padStart(2, '0') + '/' + 
                          String(fechaActual.getMonth() + 1).padStart(2, '0') + '/' + 
                          fechaActual.getFullYear();

    sheet.getRange(fila, 4).setValue(nuevaFechaTexto);
    sheet.getRange(fila, 6).setValue("Sí");

    return "✅ ¡Préstamo aplazado con éxito por 3 días más!";
  } catch (err) {
    return "❌ Error al aplazar préstamo: " + err.toString();
  }
}

function obtenerPrestamosEnServer() {
  try {
    const doc = SpreadsheetApp.openById(ID_HOJA);
    const sheetPrestamos = doc.getSheetByName("Hoja 1") || doc.getSheets()[0];
    const datos = sheetPrestamos.getDataRange().getDisplayValues();
    
    const prestamos = [];
    
    for (let i = 1; i < datos.length; i++) {
      const codigo = datos[i][0];
      const libro = datos[i][1];
      const correo = datos[i][2];
      const fecha = datos[i][3];
      const estado = datos[i][4];

      if (codigo && estado === "Prestado") {
        prestamos.push({
          fila: i + 1,
          codigo: codigo,
          libro: libro,
          correo: correo,
          fecha: fecha
        });
      }
    }
    
    return prestamos.reverse();
  } catch (err) {
    Logger.log("Error al obtener préstamos: " + err.toString());
    return [];
  }
}

function buscarLibroEnServer(codigo) {
  try {
    const doc = SpreadsheetApp.openById(ID_HOJA);
    const sheetInventario = doc.getSheetByName("Inventario") || doc.getSheets()[0];
    const datos = sheetInventario.getDataRange().getDisplayValues();

    const codigoBuscado = String(codigo).trim().toLowerCase();

    for (let i = 1; i < datos.length; i++) {
      const codigoCelda = String(datos[i][0]).trim().toLowerCase();
      if (codigoCelda === codigoBuscado) {
        return { encontrado: true, titulo: datos[i][1] };
      }
    }
    return { encontrado: false, titulo: "" };
  } catch (err) {
    return { encontrado: false, error: err.toString() };
  }
}

function registrarPrestamoEnServer(datosForm) {
  try {
    const doc = SpreadsheetApp.openById(ID_HOJA);
    const sheetPrestamos = doc.getSheetByName("Hoja 1") || doc.getSheets()[0];

    let fechaFormateada = datosForm.fecha;
    if (datosForm.fecha && datosForm.fecha.includes('-')) {
      const partes = datosForm.fecha.split('-');
      fechaFormateada = `${partes[2]}/${partes[1]}/${partes[0]}`;
    }

    sheetPrestamos.appendRow([datosForm.codigo, datosForm.libro, datosForm.correo, fechaFormateada, "Prestado", "No"]);
    return "Éxito";
  } catch (err) {
    return "Error: " + err.toString();
  }
}

function devolverLibroEnServer(fila) {
  try {
    const doc = SpreadsheetApp.openById(ID_HOJA);
    const sheetPrestamos = doc.getSheetByName("Hoja 1") || doc.getSheets()[0];
    
    const codigo = sheetPrestamos.getRange(fila, 1).getValue();
    const libro = sheetPrestamos.getRange(fila, 2).getValue();
    const correo = sheetPrestamos.getRange(fila, 3).getValue();
    const fechaProg = sheetPrestamos.getRange(fila, 4).getValue();

    sheetPrestamos.getRange(fila, 5).setValue("Devuelto");

    let sheetEntregados = doc.getSheetByName("Entregados");
    if (!sheetEntregados) {
      sheetEntregados = doc.insertSheet("Entregados");
      sheetEntregados.appendRow(["Código", "Libro", "Correo Estudiante", "Fecha Programada", "Fecha Entrega Real"]);
      sheetEntregados.getRange(1, 1, 1, 5).setFontWeight("bold");
    }

    const hoy = new Date();
    const fechaEntrega = `${String(hoy.getDate()).padStart(2, '0')}/${String(hoy.getMonth() + 1).padStart(2, '0')}/${hoy.getFullYear()}`;

    sheetEntregados.appendRow([codigo, libro, correo, fechaProg, fechaEntrega]);
    
    return "Éxito";
  } catch (err) {
    return "Error: " + err.toString();
  }
}

function enviarRecordatorios2DiasAntes() {
  const doc = SpreadsheetApp.openById(ID_HOJA);
  const sheet = doc.getSheetByName("Hoja 1") || doc.getSheets()[0];
  const datos = sheet.getDataRange().getValues();
  
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const baseUrl = ScriptApp.getService().getUrl();

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
        
        const urlAplazar = `${baseUrl}?action=aplazar&fila=${i + 1}`;
        
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
