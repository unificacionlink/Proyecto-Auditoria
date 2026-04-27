const ID_EXCEL = '1Z_rLuliAdV9A-6uEpLRogjtz4yAAk6ZmpLqkmj9e49c'; 

function doGet(e) {
  if (e.parameter.modo === 'firma') {
    // Usamos createTemplateFromFile para poder enviarle el DNI al HTML
    let template = HtmlService.createTemplateFromFile('FirmaUsuario');
    template.dniUrl = e.parameter.dni || ""; 
    return template.evaluate()
      .setTitle('Firma de Acta - Colaborador')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Generador de Actas TI')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ✨ NUEVA FUNCIÓN PARA ENVIAR EL CORREO CON EL ENLACE DE FIRMA
function enviarLinkPorCorreo(emailDestino, nombreColaborador, linkFirma) {
  try {
    const asunto = "Firma Requerida: Acta de Asignación de Equipo TI";
    
    // Diseñamos un correo bonito en HTML
    const cuerpoHtml = `
      <div style="font-family: Arial, sans-serif; padding: 20px; color: #333; max-width: 600px; border: 1px solid #ddd; border-radius: 8px;">
        <h2 style="color: #0056b3;">Hola, ${nombreColaborador}</h2>
        <p>El área de TI ha generado una nueva <b>Acta de Gestión de Equipos</b> que requiere tu revisión y firma.</p>
        <p>Por favor, ingresa al siguiente enlace desde tu celular o computadora para verificar los datos y registrar tu firma digital:</p>
        
        <div style="text-align: center; margin: 30px 0;">
          <a href="${linkFirma}" style="background-color: #0056b3; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 16px;">Revisar y Firmar Acta</a>
        </div>
        
        <p style="font-size: 14px;">Si el botón no funciona, puedes copiar y pegar el siguiente enlace en tu navegador:</p>
        <p style="font-size: 13px; color: #555; word-break: break-all;"><a href="${linkFirma}">${linkFirma}</a></p>
        
        <hr style="border: 1px solid #eee; margin-top: 30px;">
        <p style="font-size: 12px; color: #888; text-align: center;">Este es un mensaje automático del sistema de Gestión de TI. Por favor, no respondas a este correo.</p>
      </div>
    `;
    
    // Enviamos el correo usando el servicio de Google
    MailApp.sendEmail({
      to: emailDestino,
      subject: asunto,
      htmlBody: cuerpoHtml
    });
    
    return true; // Éxito
    
  } catch (error) {
    throw new Error("No se pudo enviar el correo: " + error.toString());
  }
}

// ==========================================
// FASE 1: TI GENERA EL BORRADOR
// ==========================================
function procesarFormulario(datos) {
  try {
    const ID_PC_ENTREGA = '1BTI_lZS3Fuf-Kr_bzeuJGULLhoUQ4_V7-K-WAQuijgA';
    const ID_PC_DEVOLUCION = '1mrUMaDJ5HLSubz2w20VLTbbWF3JmgqTAjq8MZvlw_H4';
    const ID_CELULAR_ENTREGA = '15RdTstqxKNtUtFsBiU_7cvts3fJkDMhe-FcY6tafZxY';
    const ID_CELULAR_DEVOLUCION = '19YWRadsyPxaTJPszBVnp1mTAWO-fR_5fcR3931R2BeU';

    // ID_EXCEL debe estar definida globalmente o ponla aquí:
    // const ID_EXCEL = 'TU_ID_DE_EXCEL_AQUI'; 

    let ID_PLANTILLA = (datos.equipoDrop === 'Computadora portatil') 
      ? (datos.motivo === 'Entrega de equipo' ? ID_PC_ENTREGA : ID_PC_DEVOLUCION)
      : (datos.motivo === 'Entrega de equipo' ? ID_CELULAR_ENTREGA : ID_CELULAR_DEVOLUCION);

    const folder = DriveApp.getRootFolder();
    const docOriginal = DriveApp.getFileById(ID_PLANTILLA);
    const nombreDoc = `Acta_${datos.nombres.replace(/ /g, '_')}_${Utilities.formatDate(new Date(), "GMT-5", "yyyyMMdd_HHmm")}`;
    
    const docCopia = docOriginal.makeCopy(nombreDoc, folder);
    const docId = docCopia.getId();
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    body.replaceText('<<FECHA>>', datos.fechaTramite || '');
    body.replaceText('<<NOMBRES>>', datos.nombres || '');
    body.replaceText('<<DNI>>', datos.dni || '');
    body.replaceText('<<EMAIL>>', datos.email || '-');
    body.replaceText('<<CARGO>>', datos.cargo || '-');
    body.replaceText('<<AREA>>', datos.area || '-');
    body.replaceText('<<MOTIVO>>', datos.motivo || '');
    body.replaceText('<<OBSERVACION>>', datos.observacion || '-');

    body.replaceText('<<CHK_PC>>', (datos.equipoDrop === 'Computadora portatil' ? '☑' : '☐'));
    body.replaceText('<<CHK_CELULAR>>', (datos.equipoDrop === 'Equipo Celular' ? '☑' : '☐'));
    body.replaceText('<<CHK_A>>', (datos.tipoEquipoAB === 'A' ? '☑' : '☐'));
    body.replaceText('<<CHK_B>>', (datos.tipoEquipoAB === 'B' ? '☑' : '☐'));

    body.replaceText('<<MARCA>>', datos.marca || '-');
    body.replaceText('<<MODELO>>', datos.modelo || '-');
    body.replaceText('<<SERIAL>>', datos.serial || '-');
    body.replaceText('<<PROCESADOR>>', datos.procesador || '-');
    body.replaceText('<<RAM>>', datos.ram || '-');
    body.replaceText('<<ALMACENAMIENTO>>', datos.almacenamiento || '-');
    
    body.replaceText('<<MONITOR>>', datos.monitor || '-');
    body.replaceText('<<MONI>>', datos.monitorSN || '-');
    body.replaceText('<<car>>', datos.cargadorMod || '-');
    body.replaceText('<<CARG>>', datos.cargadorSN || '-');

    body.replaceText('<<FIRMA_TI>>', '');
    doc.saveAndClose();

    // 4. GUARDADO MULTI-HOJA (General, Laptops, Celulares)
    try {
      const libro = SpreadsheetApp.openById(ID_EXCEL);
      
      // Preparamos la fila de datos
      const fila = [
        datos.fechaTramite, datos.nombres, datos.dni, datos.area, 
        datos.motivo, datos.equipoDrop, datos.marca, datos.modelo, 
        datos.serial, datos.procesador, datos.ram, datos.almacenamiento, 
        "", "Pendiente", docId, datos.email
      ];

      // A. Guardar siempre en General (Hoja 1)
      const hojaGeneral = libro.getSheetByName("Hoja 1");
      if (hojaGeneral) hojaGeneral.appendRow(fila);

      // B. Guardar en Laptops (Hoja 2) si corresponde
      if (datos.equipoDrop === 'Computadora portatil') {
        const hojaLaptops = libro.getSheetByName("Hoja 2");
        if (hojaLaptops) hojaLaptops.appendRow(fila);
      } 
      
      // C. Guardar en Celulares (Hoja 3) si corresponde
      else if (datos.equipoDrop === 'Equipo Celular') {
        const hojaCelulares = libro.getSheetByName("Hoja 3");
        if (hojaCelulares) hojaCelulares.appendRow(fila);
      }

    } catch(e) {
      console.log("Error al guardar en las hojas de Excel: " + e);
    }

    // 5. GENERAR EL LINK AUTOMÁTICO PARA EL COLABORADOR
    const urlScript = ScriptApp.getService().getUrl();
    const linkFirma = urlScript + "?modo=firma&dni=" + datos.dni;

    return { 
      exito: true, 
      mensaje: "Acta generada y guardada en las hojas correspondientes.",
      link: linkFirma 
    };

  } catch (error) {
    return { error: error.toString() };
  }
}

// ==========================================
// FASE 2: COLABORADOR FIRMA Y SE CIERRA PDF
// ==========================================
function buscarActaPendiente(dni) {
  try {
    const hoja = SpreadsheetApp.openById(ID_EXCEL).getActiveSheet();
    const datos = hoja.getDataRange().getValues();
    
    // Buscamos de abajo hacia arriba la última acta de ese DNI que esté Pendiente
    for (let i = datos.length - 1; i > 0; i--) {
      // DNI está en columna C (índice 2), Estado en columna N (índice 13)
      if (datos[i][2] == dni && datos[i][13] == 'Pendiente') {
        return {
          encontrado: true,
          fila: i + 1, // +1 porque los arrays empiezan en 0 y las filas en Sheets en 1
          nombres: datos[i][1],
          motivo: datos[i][4],
          equipo: datos[i][5],
          docId: datos[i][14] // ID_DOC está en columna O (índice 14)
        };
      }
    }
    return { encontrado: false };
  } catch (error) {
    return { error: error.toString() };
  }
}

function procesarFirmaResponsable(datos) {
  try {
    const docId = datos.docId;
    const docOriginal = DriveApp.getFileById(docId);
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // 1. Insertar la firma del colaborador
    insertarImagenFirma(body, '<<FIRMA_RESPONSABLE>>', datos.firmaResponsable);
    doc.saveAndClose();

    // 2. Generar PDF final
    const folder = DriveApp.getRootFolder();
    const nombrePdf = docOriginal.getName() + "_Firmado";
    const pdfBlob = docOriginal.getAs(MimeType.PDF);
    const archivoPdf = folder.createFile(pdfBlob).setName(nombrePdf + ".pdf");
    const urlPdf = archivoPdf.getUrl();

    // 3. --- ACTUALIZACIÓN MULTI-HOJA ---
    const libro = SpreadsheetApp.openById(ID_EXCEL);
    const hojasABuscar = ["Hoja 1", "Hoja 2", "Hoja 3"];
    
    let nombreColaborador = "";
    let emailColaborador = "";

    hojasABuscar.forEach(nombreDePestaña => {
      let hoja = libro.getSheetByName(nombreDePestaña);
      if (hoja) {
        let datosHoja = hoja.getDataRange().getValues();
        
        // Buscamos en qué fila está el docId (Columna O es índice 14)
        for (let i = 1; i < datosHoja.length; i++) {
          if (datosHoja[i][14] === docId) {
            
            // Si es la Hoja 1, guardamos los datos para el correo
            if (nombreDePestaña === "Hoja 1") {
              nombreColaborador = datosHoja[i][1]; // Columna B
              emailColaborador = datosHoja[i][15]; // Columna P
            }

            // Actualizamos Link (Col M / 13) y Estado (Col N / 14)
            hoja.getRange(i + 1, 13).setValue(urlPdf);
            hoja.getRange(i + 1, 14).setValue("Firmado");
            break; // Deja de buscar en esta pestaña y pasa a la siguiente
          }
        }
      }
    });

    // 4. --- ENVÍO DE CORREO ---
    if (emailColaborador && emailColaborador !== "-") {
      enviarActaFirmada(emailColaborador, nombreColaborador, pdfBlob);
    }

    // 5. Limpieza y retorno
    docOriginal.setTrashed(true);
    const pdfB64 = Utilities.base64Encode(pdfBlob.getBytes());
    return { exito: true, contenido: pdfB64, nombreArchivo: nombrePdf + '.pdf' };

  } catch(error) {
    return { error: error.toString() };
  }
}

function buscarDatosPorDNI(dni) {
  // IDs de tus dos bases de datos
  const ID_DB_PERSONAL = "123J9FsE1yJNK-YYRkwI94a9ZwurjV2Rmpyxo2xnDmqc"; // GOYDELSAC
  const ID_EXCEL_ACTAS = "1Z_rLuliAdV9A-6uEpLRogjtz4yAAk6ZmpLqkmj9e49c"; // Auditoría

  try {
    // ==========================================
    // INTENTO 1: Buscar en la base principal GOYDELSAC
    // ==========================================
    const ssPersonal = SpreadsheetApp.openById(ID_DB_PERSONAL);
    const hojaPersonal = ssPersonal.getSheetByName("GOYDELSAC"); 
    
    if (hojaPersonal) {
      const datosPersonal = hojaPersonal.getDataRange().getValues();
      for (let i = 1; i < datosPersonal.length; i++) {
        if (datosPersonal[i][0].toString().trim() === dni.toString().trim()) {
          return {
            exito: true,
            fuente: "GOYDELSAC",
            nombres: datosPersonal[i][1], // Columna B
            cargo: datosPersonal[i][2],   // Columna C
            area: datosPersonal[i][3]     // Columna D
          };
        }
      }
    }

    // ==========================================
    // INTENTO 2: Si no está en GOYDELSAC, buscar en AUDITORÍA
    // ==========================================
    const ssActas = SpreadsheetApp.openById(ID_EXCEL_ACTAS);
    const hojaActas = ssActas.getActiveSheet();
    const datosActas = hojaActas.getDataRange().getValues();
    
    // Buscamos de abajo hacia arriba para traer su acta más reciente
    for (let i = datosActas.length - 1; i > 0; i--) {
      // DNI en Auditoría está en la Columna C (índice 2)
      if (datosActas[i][2].toString().trim() === dni.toString().trim()) {
        return {
          exito: true,
          fuente: "Auditoría",
          nombres: datosActas[i][1], // Columna B
          area: datosActas[i][3],    // Columna D
          cargo: "" // Lo dejamos vacío porque en Auditoría no guardas el cargo
        };
      }
    }

    // ==========================================
    // Si no está en ninguna de las dos bases
    // ==========================================
    return { exito: false, mensaje: "DNI no encontrado en ninguna base de datos." };

  } catch (e) {
    return { exito: false, mensaje: "Error de conexión: " + e.toString() };
  }
}
// Función auxiliar para validación DNI en pantalla TI
function consultarDNI(dni) {
  try {
    const hoja = SpreadsheetApp.openById(ID_EXCEL).getActiveSheet();
    const datos = hoja.getDataRange().getValues();
    for (let i = datos.length - 1; i > 0; i--) {
      if (datos[i][2] == dni) return { encontrado: true, nombres: datos[i][1], area: datos[i][3] };
    }
    return { encontrado: false };
  } catch (e) { return { error: e.toString() }; }
}

function insertarImagenFirma(body, etiqueta, dataUrl) {
  const elements = body.findText(etiqueta);
  if (elements && dataUrl && dataUrl.includes('base64,')) {
    const textElement = elements.getElement();
    const base64Data = dataUrl.split(',')[1];
    const imageBlob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', 'firma.png');
    
    // Obtenemos el párrafo donde está la etiqueta
    const parrafo = textElement.getParent().asParagraph();
    
    // Insertamos la imagen
    const img = parrafo.insertInlineImage(0, imageBlob);
    
    // Borramos el texto de la etiqueta
    textElement.asText().setText(""); 
    
    // Ajustamos el tamaño y FORZAMOS A QUE SE CENTRE
    img.setWidth(140).setHeight(70);
    parrafo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }
}

/**
 * Envía el correo al colaborador con el link de la firma
 */
function enviarLinkPorCorreo(emailDestino, nombreColaborador, linkFirma) {
  try {
    const asunto = "Firma Requerida: Acta de Equipo TI - " + nombreColaborador;
    
    const cuerpoHtml = `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; border: 1px solid #eee; padding: 20px; border-radius: 10px;">
        <h2 style="color: #0056b3;">Hola, ${nombreColaborador}</h2>
        <p>Se ha generado un documento de asignación de equipo que requiere tu firma electrónica.</p>
        <p>Por favor, haz clic en el botón de abajo para revisar los datos y firmar:</p>
        <div style="text-align: center; margin: 30px 0;">
          <a href="${linkFirma}" style="background-color: #0056b3; color: white; padding: 15px 25px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">REVISAR Y FIRMAR AQUÍ</a>
        </div>
        <p style="font-size: 12px; color: #777;">Si el botón no funciona, copia este link: <br>${linkFirma}</p>
        <hr style="border: 0; border-top: 1px solid #eee;">
        <p style="font-size: 11px; color: #999; text-align: center;">Área de Tecnologías de la Información</p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: emailDestino,
      subject: asunto,
      htmlBody: cuerpoHtml
    });
    
    return true;
  } catch (e) {
    throw new Error("Error al enviar email: " + e.toString());
  }
}
function darPermisos() {
  // Esto solo sirve para que Google detecte que queremos enviar correos
  MailApp.sendEmail(Session.getActiveUser().getEmail(), "Prueba de Permisos", "¡Permisos otorgados correctamente!");
}

/**
 * Envía el PDF final firmado al colaborador y al área de TI
 */
function enviarActaFirmada(emailDestino, nombreColaborador, pdfBlob) {
  try {
    const asunto = "✅ Acta Firmada: " + nombreColaborador;
    const cuerpo = "Hola " + nombreColaborador + ",\n\nAdjunto encontrarás el acta de gestión de equipo debidamente firmada.\n\nSaludos,\nÁrea de TI.";
    
    MailApp.sendEmail({
      to: emailDestino,
      subject: asunto,
      body: cuerpo,
      attachments: [pdfBlob.setName("Acta_TI_" + nombreColaborador + ".pdf")]
    });
    
    return true;
  } catch (e) {
    console.error("Error enviando PDF: " + e.toString());
  }
}