# TallerYLubricentro

/**b138b341cd9bd3e2d0df24953fd04f24c6809a9d661124d387fa3292362a9290@group.calendar.google.com
 * NOTA IMPORTANTE: Variables que deben modificarse si el sistema se replica para otro cliente
 * 
 * 1. **ID de Carpeta Principal**: El ID de la carpeta en Google Drive donde se crearán las subcarpetas
 *    - Variable: `const carpetaPrincipalId = "1I3BoZoPdCwwLcDRbxi-Ez6c4XX0txsPR";` que es la carpeta "Vehíciulos" que 
 *      debemos crear en google drive dentro de la carpeta del cliente en appsheet *    - Descripción: Reemplazar con el ID de la nueva carpeta principal del cliente *  
 *       en      Google Drive.
 * 
 * 2. **ID del Calendario de Google**: El ID del calendario donde se registrarán los turnos
 *    - Variable: `const calendarId = "b138b341cd9bd3e2d0df24953fd04f24c6809a9d661124d387fa3292362a9290@group.calendar.google.com";`
 *    - Descripción: Reemplazar con el ID del nuevo calendario de Google del cliente.
 * 
 * 3. **Datos del Taller**: Información del taller que se encuentra en la hoja "Datos Taller"
 *    - Variables: 
 *      - `const nombreTaller = hojaDatosTaller.getRange("B2").getValue();`
 *      - `const direccionTaller = hojaDatosTaller.getRange("B3").getValue();`
 *      - `const telefonoTaller = hojaDatosTaller.getRange("B4").getValue();`
 *      - `const emailTaller = hojaDatosTaller.getRange("B5").getValue();`
 *      - `const whatsappTaller = hojaDatosTaller.getRange("B6").getValue();`
 *      - `const nombreGerente = hojaDatosTaller.getRange("B7").getValue();`
 *    - Descripción: Asegúrate de que estos datos se configuren correctamente en la hoja "Datos Taller" para el nuevo cliente.
 * 
 * 4. **URLs de Generación de QR**: La URL que se usa para generar códigos QR
 *    - Variable: `const urlQR = "https://quickchart.io/qr?text=${encodeURIComponent(datoParaQR)}&size=150";`
 *    - Descripción: Esta URL debería permanecer igual a menos que quieras cambiar el servicio de generación de QR.
 * 
 * Recuerda buscar y reemplazar estos elementos cuando configures el sistema para un nuevo cliente.
 */


function generarMapaDePlanilla() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = ss.getSheets();
  const mapa = [];

  hojas.forEach((hoja) => {
    const nombreHoja = hoja.getName();
    const rangoDatos = hoja.getDataRange();
    const encabezados = rangoDatos.getValues()[0];
    const totalFilas = rangoDatos.getNumRows();
    const totalColumnas = rangoDatos.getNumColumns();

    mapa.push({
      nombreHoja: nombreHoja,
      encabezados: encabezados,
      totalFilas: totalFilas,
      totalColumnas: totalColumnas
    });
  });

  Logger.log(JSON.stringify(mapa, null, 2));
}

// hasta quó OK

// Funciones auxiliares de formateo
function formatearNombre(nombre) {
  return nombre.toLowerCase().replace(/\b\w/g, l => l.toUpperCase());
}

function formatearPatente(patente) {
  return patente.toUpperCase().replace(/\s+/g, '');
}

function formatearOracion(texto) {
  texto = texto.toLowerCase().trim();
  return texto.charAt(0).toUpperCase() + texto.slice(1);
}

function estandarizaciondetextos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = []; // Inicializar un registro para almacenar los cambios realizados

  // Validar día y horario laboral
  const ahora = new Date();
  const dia = ahora.getDay(); // 0: Domingo, 1: Lunes, ..., 6: Sábado
  const hora = ahora.getHours(); // Hora actual (0 a 23)

  if (dia === 0 || dia === 6 || hora < 9 || hora >= 18) {
    Logger.log("Fuera del horario laboral (lunes a viernes, 9:00 a 18:00). La función no se ejecuta.");
    return;
  }

  // Función auxiliar para procesar y estandarizar una hoja específica
  function procesarHoja(hoja, columnas) {
    const totalFilas = hoja.getLastRow() - 1;
    if (totalFilas <= 0) return; // No hay filas para procesar

    const rangoDatos = hoja.getRange(2, 1, totalFilas, hoja.getLastColumn());
    const datos = rangoDatos.getValues();

    datos.forEach((fila, index) => {
      columnas.forEach(columna => {
        if (fila[columna.index] && typeof fila[columna.index] === 'string') {
          const valorOriginal = fila[columna.index];
          fila[columna.index] = columna.formato(valorOriginal);
          if (valorOriginal !== fila[columna.index]) {
            log.push(`${hoja.getName()} - Fila ${index + 2}, Columna ${columna.nombre}: "${valorOriginal}" => "${fila[columna.index]}"`);
          }
        }
      });
    });

    rangoDatos.setValues(datos); // Aplicar los cambios de una sola vez
  }

  // Configuración para cada hoja
  const configuraciones = [
    {
      hoja: ss.getSheetByName("Base de Clientes"),
      columnas: [
        { index: 1, nombre: "Nombre", formato: formatearNombre },
        { index: 2, nombre: "Empresa", formato: formatearOracion },
        { index: 3, nombre: "Dirección", formato: formatearOracion },
        { index: 4, nombre: "Dirección Google Maps", formato: formatearOracion },
        { index: 8, nombre: "Email", formato: (email) => email.toLowerCase() }
      ]
    },
    {
      hoja: ss.getSheetByName("Base Vehículos"),
      columnas: [
        { index: 1, nombre: "Nombre", formato: formatearNombre },
        { index: 2, nombre: "Patente", formato: formatearPatente },
        { index: 3, nombre: "Modelo", formato: formatearOracion }
      ]
    },
    {
      hoja: ss.getSheetByName("Trabajos Hechos"),
      columnas: [
        { index: 1, nombre: "Patente", formato: formatearPatente },
        { index: 2, nombre: "Nombre", formato: formatearNombre }
      ]
    }
  ];

  // Ejecutar el proceso de estandarización para cada hoja configurada
  configuraciones.forEach(config => {
    if (config.hoja) {
      procesarHoja(config.hoja, config.columnas);
    } else {
      Logger.log(`No se encontró la hoja: ${config.hoja}`);
    }
  });

  // Mostrar el log de cambios realizados
  if (log.length > 0) {
    Logger.log("Cambios realizados:");
    log.forEach(cambio => Logger.log(cambio));
  } else {
    Logger.log("No se realizaron cambios.");
  }
    // Pausar antes de ejecutar la creación de carpetas (por ejemplo, 3 segundos)
  Utilities.sleep(3000);
  crearYRegistrarCarpetasParaVehiculos(); // Llamar a la función de creación y registro de carpetas

}


/**
 * Función principal para envío y registro de presupuestos.
 */
function envioYRegistroDePresupuesto() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaModeloPresupuesto = ss.getSheetByName("Modelo Presupuesto");
    const hojaPresupuestos = ss.getSheetByName("Presupuestos");
    const hojaDatosTaller = ss.getSheetByName("Datos Taller");

    // Validar hojas existentes
    if (!hojaModeloPresupuesto || !hojaPresupuestos || !hojaDatosTaller) {
      throw new Error("No se encontraron todas las hojas necesarias.");
    }

    // Obtener datos del presupuesto
    const cliente = hojaModeloPresupuesto.getRange("B12").getValue();
    const emailCliente = hojaModeloPresupuesto.getRange("B13").getValue();
    const patente = hojaModeloPresupuesto.getRange("D12").getValue();
    const monto = hojaModeloPresupuesto.getRange("G28").getValue() || 0;
    const fechaEnvio = hojaModeloPresupuesto.getRange("B9").getValue();
    const fechaVencimiento = hojaModeloPresupuesto.getRange("F15").getValue();
    const nombrePDF = hojaModeloPresupuesto.getRange("F12").getValue();

    if (!cliente || !patente || !nombrePDF) {
      throw new Error("Faltan datos críticos para el presupuesto. Verifique los campos.");
    }

    // Obtener datos del taller
    const nombreTaller = hojaDatosTaller.getRange("B2").getValue();
    const direccionTaller = hojaDatosTaller.getRange("B3").getValue();
    const telefonoTaller = hojaDatosTaller.getRange("B4").getValue();
    const emailTaller = hojaDatosTaller.getRange("B5").getValue();
    const whatsappTaller = hojaDatosTaller.getRange("B6").getValue();
    const nombreGerente = hojaDatosTaller.getRange("B7").getValue();

    // Confirmar envío
    const ui = SpreadsheetApp.getUi();
    const respuesta = ui.alert(
      'Confirmación',
      `Está a punto de enviar un presupuesto a "${cliente}". ¿Está seguro?`,
      ui.ButtonSet.YES_NO
    );

    if (respuesta !== ui.Button.YES) {
      Logger.log("Acción cancelada por el usuario.");
      return;
    }

    // Generar el código QR usando QuickChart
    const datoParaQR = patente.trim();
    if (!datoParaQR) {
      throw new Error("El dato en la celda D12 (Patente) está vacío. No se puede generar el QR.");
    }

    const urlQR = `https://quickchart.io/qr?text=${encodeURIComponent(datoParaQR)}&size=150`;
    const imagenQR = UrlFetchApp.fetch(urlQR).getBlob();
    hojaModeloPresupuesto.getRange("E4").clearContent();
    hojaModeloPresupuesto.insertImage(imagenQR, 5, 4).setAnchorCell(hojaModeloPresupuesto.getRange("E4"));

    // Generar carpeta de presupuestos si no existe
    const carpetaPrincipalId = "1I3BoZoPdCwwLcDRbxi-Ez6c4XX0txsPR";
    let carpetaPresupuestos = DriveApp.getFolderById(carpetaPrincipalId).getFoldersByName("Presupuestos");

    if (!carpetaPresupuestos.hasNext()) {
      carpetaPresupuestos = DriveApp.getFolderById(carpetaPrincipalId).createFolder("Presupuestos");
    } else {
      carpetaPresupuestos = carpetaPresupuestos.next();
    }

    // Generar el PDF
    const pdfBlob = generarPDF(hojaModeloPresupuesto, nombrePDF);
    const pdfFile = carpetaPresupuestos.createFile(pdfBlob).setName(`${nombrePDF}.pdf`);
    const linkPDF = pdfFile.getUrl();

    // Registrar el presupuesto
    const ultimaFila = hojaPresupuestos.getLastRow();
    const nuevoIdPresupuesto = "PRESU" + Utilities.formatString("%04d", ultimaFila);

    hojaPresupuestos.appendRow([
      nuevoIdPresupuesto, patente || "0", cliente, linkPDF, monto, fechaEnvio, "PRESUPUESTADO"
    ]);

    // Preparar y enviar el correo electrónico
    const asunto = `Presupuesto - ${nombrePDF}`;
    const cuerpo = `
      <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
        <h2 style="color: #0056b3;">Estimad@ ${cliente},</h2>
        <p>Le enviamos el presupuesto solicitado para su vehículo con patente <strong>${patente}</strong>.</p>
        <p>
          Este presupuesto tiene una validez de <strong>7 días</strong>, 
          venciendo el día <strong>${fechaVencimiento}</strong>.
        </p>
        <div style="background-color: #f9f9f9; border: 1px solid #ddd; padding: 15px; margin: 15px 0; border-radius: 5px;">
          <h4 style="margin-top: 0; color: #0056b3;">Información de contacto</h4>
          <p>
            Teléfono: <a href="tel:${telefonoTaller}" style="color: #0056b3;">${telefonoTaller}</a><br>
            WhatsApp: <a href="https://wa.me/549${whatsappTaller}" style="color: #0056b3;">${whatsappTaller}</a><br>
            Email: <a href="mailto:${emailTaller}" style="color: #0056b3;">${emailTaller}</a><br>
            Dirección: ${direccionTaller}
          </p>
        </div>
        <p>Si tiene alguna consulta, no dude en ponerse en contacto con nosotros. Estamos aquí para ayudarle.</p>
        <p style="margin-top: 30px;">Saludos cordiales,</p>
        <p>
          <strong>${nombreGerente}</strong><br>
          ${nombreTaller}
        </p>
      </div>
    `;

    if (emailCliente) {
      MailApp.sendEmail({
        to: emailCliente,
        subject: asunto,
        htmlBody: cuerpo,
        attachments: [pdfFile.getAs(MimeType.PDF)],
        name: nombreTaller
      });
      ui.alert(`Se envió el presupuesto al mail "${emailCliente}"`);
    } else {
      ui.alert("No se encontró un correo electrónico del cliente. Presupuesto registrado en la carpeta 'Presupuestos'.");
    }

    Logger.log("Presupuesto enviado o registrado exitosamente.");
  } catch (error) {
    Logger.log("Error en envioYRegistroDePresupuesto: " + error.message);
    SpreadsheetApp.getUi().alert("Se produjo un error: " + error.message);
  }
}


// Función para generar el PDF
function generarPDF(hoja, nombrePDF) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaId = hoja.getSheetId();
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${hojaId}`;

  const response = UrlFetchApp.fetch(url, {
    method: "GET",
    headers: {
      "Authorization": `Bearer ${ScriptApp.getOAuthToken()}`
    }
  });

  return response.getBlob().setName(`${nombrePDF}.pdf`);
}


// aca iniciamos con la creación y registro de carpetas por vehículos

function crearYRegistrarCarpetasParaVehiculos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaVehiculos = ss.getSheetByName("Base Vehículos");
    const hojaCarpetas = ss.getSheetByName("Carpetas");
    const carpetaPrincipalId = "1I3BoZoPdCwwLcDRbxi-Ez6c4XX0txsPR"; // Reemplazá con el ID de tu carpeta principal en Google Drive

    const carpetaPrincipal = DriveApp.getFolderById(carpetaPrincipalId);
    const carpetasData = hojaCarpetas.getRange("A2:A" + hojaCarpetas.getLastRow()).getValues();
    const patentesExistentes = new Set(carpetasData.flat());

    const vehiculosData = hojaVehiculos.getRange("A2:L" + hojaVehiculos.getLastRow()).getValues();
    const nuevasCarpetas = [];

    vehiculosData.forEach((fila) => {
      const idVehiculo = fila[0];
      const patente = fila[2];

      if (patente && !patentesExistentes.has(patente)) {
        // Crear una nueva carpeta para la patente
        const carpetaVehiculo = carpetaPrincipal.createFolder(patente);
        const urlCarpeta = carpetaVehiculo.getUrl();

        // Agregar la nueva carpeta a la lista para registrar
        nuevasCarpetas.push([patente, idVehiculo, urlCarpeta]);
        patentesExistentes.add(patente);
        Logger.log(`Carpeta creada para el vehículo con patente ${patente}. URL: ${urlCarpeta}`);
      } else {
        Logger.log(`Carpeta ya existente para la patente: ${patente}, omitiendo creación.`);
      }
    });

    if (nuevasCarpetas.length > 0) {
      hojaCarpetas.getRange(hojaCarpetas.getLastRow() + 1, 1, nuevasCarpetas.length, 3).setValues(nuevasCarpetas);
      Logger.log("Nuevas carpetas registradas en la hoja 'Carpetas'.");
    } else {
      Logger.log("No se crearon nuevas carpetas en esta ejecución.");
    }

    Logger.log("Proceso de creación y registro de carpetas completado.");
  } catch (error) {
    Logger.log("Error en crearYRegistrarCarpetasParaVehiculos: " + error.message);
    SpreadsheetApp.getUi().alert("Se produjo un error: " + error.message);
  }
}

// hasta la fila 342 - todo perfeto - ahora seguimos con los turnos

/**
 * Función para gestionar y sincronizar los turnos con Google Calendar.
 * Registra el evento en el calendario y agrega al cliente como participante
 * sin enviar invitaciones o notificaciones por correo electrónico.
 */

function gestionarTurnosEnCalendario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTurnos = ss.getSheetByName("Turnos");
  const hojaClientes = ss.getSheetByName("Base de Clientes");
  const hojaDatosTaller = ss.getSheetByName("Datos Taller");
  let log = []; // Registro de acciones realizadas

  // Obtener datos del taller (ajustando por los encabezados en la primera fila)
  const datosTaller = hojaDatosTaller.getRange(2, 2, hojaDatosTaller.getLastRow() - 1).getValues().flat();
  const [nombreTaller, direccion, telefono, emailTaller, whatsapp, nombreGerente] = datosTaller;

  // Configurar el calendario
  const calendarId = "b138b341cd9bd3e2d0df24953fd04f24c6809a9d661124d387fa3292362a9290@group.calendar.google.com";
  const calendar = CalendarApp.getCalendarById(calendarId);

  // Obtener los datos de la hoja "Turnos", omitiendo la fila de encabezados
  const rangoTurnos = hojaTurnos.getRange(2, 1, hojaTurnos.getLastRow() - 1, hojaTurnos.getLastColumn());
  const datosTurnos = rangoTurnos.getValues();

  datosTurnos.forEach((fila, index) => {
    const [idTurno, patente, cliente, fecha, hora, motivo, estado, idCalendar] = fila;

    // Validar datos esenciales
    if (!cliente || !(fecha instanceof Date) || !hora) {
      log.push(`Fila ${index + 2}: Error - Datos incompletos o malformateados.`);
      return;
    }

    // Procesar la hora y asegurarse de que sea un objeto Date válido
    let horas, minutos;
    if (typeof hora === 'string' && hora.includes(":")) {
      [horas, minutos] = hora.split(":").map(Number);
    } else if (hora instanceof Date) {
      horas = hora.getHours();
      minutos = hora.getMinutes();
    } else {
      log.push(`Fila ${index + 2}: Error - Hora inválida o malformateada.`);
      return;
    }

    // Crear el objeto Date para la hora del turno
    fecha.setHours(horas, minutos);
    const fechaFin = new Date(fecha.getTime() + 60 * 60 * 1000); // Duración del evento: 1 hora

    // Verificar si el evento ya existe
    if (idCalendar) {
      const eventoExistente = calendar.getEventById(idCalendar);
      if (eventoExistente) {
        log.push(`Fila ${index + 2}: Evento ya existe en el calendario con ID: ${idCalendar}`);
        return;
      }
    }

    // Crear el nombre del evento con el formato: "NombreTaller - Cliente - Patente"
    const nombreEvento = `${nombreTaller} - ${cliente} - ${patente || ""}`.trim();

    // Crear el evento en Google Calendar
    try {
      // Obtener el email del cliente
      const emailCliente = buscarEmailCliente(cliente, hojaClientes);

      // Configurar opciones del evento
      const opcionesEvento = {
        location: direccion,
        description: `${motivo}\n\n Te pedimos que seas puntual para aprovechar al máximo el tiempo de tu turno y evitar demoras a otros clientes. ¡Muchas gracias, te esperamos!`,
        guests: emailCliente || "", // Si hay email, se agrega; si no, queda vacío
        sendInvites: !!emailCliente, // Envía la invitación solo si hay email del cliente
        conferencing: { useDefault: false } // Desactiva la videoconferencia
      };

      // Crear el evento en el calendario
      const evento = calendar.createEvent(nombreEvento, fecha, fechaFin, opcionesEvento);
      fila[7] = evento.getId(); // Guardar el ID del evento en la columna "ID Calendar"
      log.push(`Fila ${index + 2}: Evento creado y sincronizado en el calendario. ID de evento: ${evento.getId()}`);
    } catch (error) {
      log.push(`Fila ${index + 2}: Error al crear el evento - ${error.message}`);
    }
  });

  // Guardar los datos actualizados en la hoja "Turnos"
  rangoTurnos.setValues(datosTurnos);

  // Mostrar el log de acciones realizadas
  if (log.length > 0) {
    Logger.log("Registro de acciones:");
    log.forEach(accion => Logger.log(accion));
  } else {
    Logger.log("No se realizaron cambios o no se detectaron errores.");
  }
}

// Función para buscar el email del cliente en la hoja "Base de Clientes"
function buscarEmailCliente(cliente, hojaClientes) {
  const datosClientes = hojaClientes.getDataRange().getValues();
  for (let i = 1; i < datosClientes.length; i++) {
    if (datosClientes[i][1] === cliente) { // Asume que la columna B es el nombre del cliente
      return datosClientes[i][8]; // Asume que la columna I contiene el email del cliente
    }
  }
  return null;
}


/**
 * Función para buscar el email del cliente en la hoja "Base de Clientes".
 * @param {string} nombreCliente - El nombre del cliente a buscar.
 * @param {Sheet} hojaClientes - La hoja "Base de Clientes".
 * @returns {string|null} - El email del cliente, o null si no se encuentra.
 */
function buscarEmailCliente(nombreCliente, hojaClientes) {
  const datosClientes = hojaClientes.getDataRange().getValues();
  for (let i = 1; i < datosClientes.length; i++) {
    if (datosClientes[i][1] === nombreCliente) { // Columna B - Nombre del cliente
      return datosClientes[i][8]; // Columna I - Email
    }
  }
  return null;
}

// linea 543 - todo ok hasta acá

/**
 * Función para actualizar la deuda en la columna I de "Trabajos Hechos".
 * Calcula la diferencia entre "Monto" (columna K) y "Pagó" (columna J).
 */
function actualizarDeudaTrabajos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTrabajos = ss.getSheetByName("Trabajos Hechos");
  
  // Obtener los datos de la hoja "Trabajos Hechos" omitiendo la fila de encabezados
  const datos = hojaTrabajos.getRange(2, 1, hojaTrabajos.getLastRow() - 1, hojaTrabajos.getLastColumn()).getValues();
  
  // Array para almacenar filas actualizadas
  const datosActualizados = datos.map((fila) => {
    const monto = parseFloat(fila[9]) || 0; // Columna K (Monto)
    const pago = parseFloat(fila[10]) || 0; // Columna J (Pagó)
    const diferencia = monto - pago;
    
    // Actualizar la columna I (Debe) con la diferencia y columna L con el estado
    fila[8] = diferencia > 0 ? diferencia : 0; // Deuda solo si es mayor a cero
    fila[11] = diferencia > 0 ? "Debe" : "Pago Completo"; // Estado
    
    return fila;
  });
  
  // Escribir los datos actualizados de vuelta en "Trabajos Hechos"
  hojaTrabajos.getRange(2, 1, datosActualizados.length, datosActualizados[0].length).setValues(datosActualizados);
  Logger.log("Deudas actualizadas en 'Trabajos Hechos'");
}

/**
 * Función para copiar la estructura de "Trabajos Hechos" y
 * solo las líneas donde el dato de la columna H es "03_LISTO PARA ENTREGAR" o "04_ENTREGADO"
 * y la diferencia entre "Monto" (columna K) y "Pagó" (columna J) es mayor que cero.
 * Coloca la diferencia en la columna I de "Reporte Deuda".
 */

/**
 * Función para generar el "Reporte Deuda" basado en los datos actuales de "Trabajos Hechos".
 */
function generarReporteDeuda() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTrabajos = ss.getSheetByName("Trabajos Hechos");
  let hojaReporte = ss.getSheetByName("Reporte Deuda");

  // Crear la hoja "Reporte Deuda" si no existe
  if (!hojaReporte) {
    hojaReporte = ss.insertSheet("Reporte Deuda");
  } else {
    hojaReporte.clear(); // Limpiar contenido si la hoja ya existe
  }

  // Copiar los encabezados de la hoja "Trabajos Hechos"
  const encabezados = hojaTrabajos.getRange(1, 1, 1, hojaTrabajos.getLastColumn()).getValues();
  hojaReporte.getRange(1, 1, 1, encabezados[0].length).setValues(encabezados);

  // Obtener los datos de la hoja "Trabajos Hechos" omitiendo la fila de encabezados
  const datos = hojaTrabajos.getRange(2, 1, hojaTrabajos.getLastRow() - 1, hojaTrabajos.getLastColumn()).getValues();

  // Filtrar filas donde la deuda en la columna I es mayor que cero y el estado en la columna H es "03_LISTO PARA ENTREGAR" o "04_ENTREGADO"
  const filasConDeuda = datos.filter(fila => 
    parseFloat(fila[8]) > 0 && 
    (fila[7] === "03_LISTO PARA ENTREGAR" || fila[7] === "04_ENTREGADO")
  );

  // Copiar las filas filtradas a la hoja "Reporte Deuda"
  if (filasConDeuda.length > 0) {
    hojaReporte.getRange(2, 1, filasConDeuda.length, filasConDeuda[0].length).setValues(filasConDeuda);
  }

  Logger.log(`Se generó el "Reporte Deuda" con ${filasConDeuda.length} líneas.`);
}


//resuelto el informe de deudas - vamos a los informes de los trabajos

/**
 * Función principal para generar informes de "Trabajos Hechos".
 * Crea informes en formato PDF, genera códigos QR, guarda en carpetas y envía correos si se dispone de email.
 */
function generarInformeDeTrabajosHechos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTrabajos = ss.getSheetByName("Trabajos Hechos");
  const hojaInforme = ss.getSheetByName("Informe Trabajo");
  const hojaClientes = ss.getSheetByName("Base de Clientes");
  const hojaVehiculos = ss.getSheetByName("Base Vehículos");
  const hojaDatosTaller = ss.getSheetByName("Datos Taller");
  const carpetaPrincipalId = "1I3BoZoPdCwwLcDRbxi-Ez6c4XX0txsPR";

  // Verificar que la hoja "Informe Trabajo" existe
  if (!hojaInforme) {
    Logger.log("Error: La hoja 'Informe Trabajo' no se encontró.");
    return;
  }

  // Mostrar la hoja "Informe Trabajo" si está oculta
  const hojaOculta = hojaInforme.isSheetHidden();
  if (hojaOculta) hojaInforme.showSheet();

  // Obtener los datos de "Trabajos Hechos"
  const trabajosData = hojaTrabajos.getDataRange().getValues();
  Logger.log(`Iniciando la generación de informes para ${trabajosData.length - 1} trabajos.`);

  // Recorrer cada fila de la hoja "Trabajos Hechos"
  trabajosData.forEach((fila, index) => {
    if (index === 0) {
      Logger.log("Saltando encabezado de la tabla.");
      return; // Saltar encabezados
    }

    const [idTrabajo, patente, cliente, motivo, detalle, fechaIngreso, , estado, , montoPresupuestado, montoPagado, , linkInforme] = fila;
    const montoRestante = montoPresupuestado - montoPagado;

    // Verificar si ya existe un link de informe en la columna M
    if (linkInforme) {
      Logger.log(`Informe ya existente para el vehículo ${patente} en la fila ${index + 1}. Omitiendo creación.`);
      return;
    }

// Verificar estado y validez de datos
if ((estado === "03_LISTO PARA ENTREGAR" || estado === "04_ENTREGADO") && patente && cliente) {
  Logger.log(`Procesando trabajo en la fila ${index + 1} para el vehículo con patente ${patente}.`);

  const fechaActual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
  const nombreClienteSanitizado = cliente.replace(/[^a-zA-Z0-9]/g, "_");
  const nombreInforme = `Informe_Trabajo_${nombreClienteSanitizado}_${patente}_${fechaActual}`;
  Logger.log(`Nombre del informe generado: ${nombreInforme}`);

  // Buscar o crear la carpeta del vehículo
  const carpetaVehiculo = buscarOCrearCarpetaVehiculo(patente, carpetaPrincipalId);
  if (!carpetaVehiculo) {
    Logger.log(`Error: No se pudo encontrar o crear la carpeta para el vehículo ${patente}.`);
    return;
  }

      // Verificar si el informe ya existe en Drive
      const archivos = carpetaVehiculo.getFilesByName(nombreInforme);
      if (archivos.hasNext()) {
        Logger.log(`Informe ya existente en Drive para el vehículo ${patente}. Omitiendo creación.`);
        return;
      }

      try {
        // Completar la hoja "Informe Trabajo" con los datos del trabajo
        Logger.log(`Completando datos en la hoja 'Informe Trabajo' para el vehículo ${patente}.`);
        hojaInforme.getRange("B12").setValue(cliente);
        hojaInforme.getRange("D12").setValue(patente);
        hojaInforme.getRange("F12").setValue(idTrabajo);
        hojaInforme.getRange("F15").setValue(new Date());

        const emailCliente = obtenerEmailCliente(cliente, hojaClientes);
        const telefonoCliente = obtenerTelefonoCliente(cliente, hojaClientes);
        const marcaVehiculo = obtenerMarcaVehiculo(patente, hojaVehiculos);
        const modeloVehiculo = obtenerModeloVehiculo(patente, hojaVehiculos);

        hojaInforme.getRange("B13").setValue(telefonoCliente);
        hojaInforme.getRange("B14").setValue(emailCliente);
        hojaInforme.getRange("D13").setValue(marcaVehiculo);
        hojaInforme.getRange("D14").setValue(modeloVehiculo);

        const fechaIngresoFormateada = Utilities.formatDate(new Date(fechaIngreso), Session.getScriptTimeZone(), "dd/MM/yyyy");
        hojaInforme.getRange("B17").setValue(
          `Fecha de Ingreso: ${fechaIngresoFormateada}\nMotivo: ${motivo}\nDetalle: ${detalle}\nMonto Presupuestado: ${montoPresupuestado}\nMonto Restante: ${montoRestante}`
        );

        hojaInforme.getRange("G26").setValue(montoPresupuestado);
        hojaInforme.getRange("G27").setValue(montoPagado);
        hojaInforme.getRange("G28").setValue(montoRestante);

        // Generar el código QR y agregarlo a la celda E4
        const datoParaQR = patente;
        const urlQR = `https://quickchart.io/qr?text=${encodeURIComponent(datoParaQR)}&size=150`;
        const imagenQR = UrlFetchApp.fetch(urlQR).getBlob();
        hojaInforme.getRange("E4").clearContent();
        hojaInforme.insertImage(imagenQR, 5, 4).setAnchorCell(hojaInforme.getRange("E4"));

        Utilities.sleep(3000); // Esperar 3 segundos

        // Confirmar que los datos se completaron correctamente
        if (hojaInforme.getRange("D12").getValue() === patente && hojaInforme.getRange("B12").getValue() === cliente) {
          Logger.log(`Datos confirmados en el template para el vehículo ${patente}. Generando PDF...`);

          // Generar el PDF y guardar en la carpeta correspondiente
          const pdfBlob = generarPDFConReintento(hojaInforme, nombreInforme);
          if (pdfBlob) {
            const archivoPDF = carpetaVehiculo.createFile(pdfBlob);
            Logger.log(`Informe PDF creado y guardado correctamente para el vehículo ${patente}.`);

            // Registrar el link del informe en la columna M de la hoja "Trabajos Hechos"
            hojaTrabajos.getRange(index + 1, 13).setValue(archivoPDF.getUrl());

            // Enviar el informe por email si el email es válido
            if (emailCliente && validarEmail(emailCliente)) {
              enviarEmailCliente(emailCliente, cliente, patente, archivoPDF, hojaDatosTaller);
              Logger.log(`Informe enviado por email a ${cliente} (${emailCliente}).`);
            } else {
              Logger.log(`Email inválido o no encontrado para el cliente ${cliente}.`);
            }
          }

          Utilities.sleep(3000); // Esperar 3 segundos
          limpiarCeldasInforme(hojaInforme); // Limpiar la hoja de informe
          Logger.log(`Informe limpiado para el siguiente trabajo.`);
        } else {
          Logger.log(`Error: Datos no confirmados en el template para el vehículo ${patente}.`);
        }
      } catch (error) {
        Logger.log(`Error al completar datos en la hoja de informe para el vehículo ${patente}: ${error.message}`);
      }
    } else {
      Logger.log(`Fila ${index + 1} omitida. Estado: ${estado}, Patente: ${patente}, Cliente: ${cliente}.`);
    }
  });

  // Ocultar la hoja "Informe Trabajo" si estaba oculta anteriormente
  if (hojaOculta) hojaInforme.hideSheet();
  Logger.log("Proceso de generación de informes completado.");
}


/**
 * Función auxiliar para buscar o crear una carpeta específica para un vehículo basado en su patente.
 */
function buscarOCrearCarpetaVehiculo(patente, carpetaPrincipalId) {
  const carpetaPrincipal = DriveApp.getFolderById(carpetaPrincipalId);
  const carpetas = carpetaPrincipal.getFoldersByName(patente);

  if (carpetas.hasNext()) {
    return carpetas.next();
  } else {
    return carpetaPrincipal.createFolder(patente);
  }
}

/**
 * Función auxiliar para obtener el email del cliente desde la hoja "Base de Clientes"
 */
function obtenerEmailCliente(cliente, hojaClientes) {
  const clientesData = hojaClientes.getDataRange().getValues();
  for (const fila of clientesData) if (fila[1] === cliente) return fila[8];
  return null;
}

/**
 * Función auxiliar para obtener el teléfono del cliente desde la hoja "Base de Clientes"
 */
function obtenerTelefonoCliente(cliente, hojaClientes) {
  const clientesData = hojaClientes.getDataRange().getValues();
  for (const fila of clientesData) if (fila[1] === cliente) return fila[6]; // Asumimos que la columna 7 contiene el teléfono
  return null;
}

/**
 * Función auxiliar para obtener la marca del vehículo desde la hoja "Base Vehículos"
 */
function obtenerMarcaVehiculo(patente, hojaVehiculos) {
  const vehiculosData = hojaVehiculos.getDataRange().getValues();
  for (const fila of vehiculosData) if (fila[1] === patente) return fila[3]; // Columna D para la marca
  return null;
}

/**
 * Función auxiliar para obtener el modelo del vehículo desde la hoja "Base Vehículos"
 */
function obtenerModeloVehiculo(patente, hojaVehiculos) {
  const vehiculosData = hojaVehiculos.getDataRange().getValues();
  for (const fila of vehiculosData) if (fila[1] === patente) return fila[4]; // Columna E para el modelo
  return null;
}

/**
 * Función auxiliar para validar el email
 */
function validarEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

/**
 * Función auxiliar para enviar el email al cliente
 */
function enviarEmailCliente(email, cliente, patente, archivoPDF, hojaDatosTaller) {
  // Obtener datos del taller
  const nombreTaller = hojaDatosTaller.getRange("B2").getValue();
  const direccionTaller = hojaDatosTaller.getRange("B3").getValue();
  const telefonoTaller = hojaDatosTaller.getRange("B4").getValue();
  const emailTaller = hojaDatosTaller.getRange("B5").getValue();
  const whatsappTaller = hojaDatosTaller.getRange("B6").getValue();

  // Configurar asunto y cuerpo del correo
  const asunto = `Informe de Trabajo - Vehículo ${patente}`;
  const cuerpo = `
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
      <h2 style="color: #0056b3;">Estimad@ ${cliente},</h2>
      <p>Le enviamos el informe del trabajo realizado en su vehículo con patente <strong>${patente}</strong>.</p>
      <p>
        En el informe encontrará todos los detalles del trabajo realizado. Si tiene alguna duda o consulta, 
        no dude en comunicarse con nosotros. Nuestro equipo estará encantado de atenderle.
      </p>
      <div style="background-color: #f9f9f9; border: 1px solid #ddd; padding: 15px; margin: 15px 0; border-radius: 5px;">
        <h4 style="margin-top: 0; color: #0056b3;">Información de contacto</h4>
        <p>
          Dirección: ${direccionTaller}<br>
          Teléfono: <a href="tel:${telefonoTaller}" style="color: #0056b3;">${telefonoTaller}</a><br>
          WhatsApp: <a href="https://wa.me/549${whatsappTaller}" style="color: #0056b3;">${whatsappTaller}</a><br>
          Email: <a href="mailto:${emailTaller}" style="color: #0056b3;">${emailTaller}</a>
        </p>
      </div>
      <p>Esperamos que este informe sea de su utilidad. Agradecemos su confianza en ${nombreTaller}.</p>
      <p style="margin-top: 30px;">Saludos cordiales,</p>
      <p>
        <strong>Equipo de ${nombreTaller}</strong>
      </p>
    </div>
  `;

  // Enviar el correo con el archivo adjunto
  MailApp.sendEmail({
    to: email,
    subject: asunto,
    htmlBody: cuerpo,
    attachments: [archivoPDF],
    name: nombreTaller // El remitente aparece como el nombre del taller
  });

  Logger.log(`Correo enviado a ${email} con el informe del vehículo ${patente}.`);
}

/**
 * Función auxiliar para generar el PDF con reintentos
 */
function generarPDFConReintento(hoja, nombrePDF) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaId = hoja.getSheetId();
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${hojaId}`;

  let pdfBlob = null;
  for (let intento = 0; intento < 3; intento++) {
    try {
      const response = UrlFetchApp.fetch(url, {
        method: "GET",
        headers: { "Authorization": `Bearer ${ScriptApp.getOAuthToken()}` },
        muteHttpExceptions: true
      });
      if (response.getResponseCode() === 200) {
        pdfBlob = response.getBlob().setName(`${nombrePDF}.pdf`);
        break;
      }
    } catch (error) {
      Logger.log(`Error en intento ${intento + 1} al generar PDF: ${error.message}`);
      Utilities.sleep(5000); // Esperar 5 segundos antes de reintentar
    }
  }

  if (!pdfBlob) Logger.log("Error: No se pudo generar el PDF después de varios intentos.");
  return pdfBlob;
}

/**
 * Función auxiliar para limpiar las celdas de la hoja "Informe Trabajo"
 */
function limpiarCeldasInforme(hojaInforme) {
  hojaInforme.getRange("B12:D14").clearContent();
  hojaInforme.getRange("F12:F15").clearContent();
  hojaInforme.getRange("B17").clearContent();
  hojaInforme.getRange("G26:G28").clearContent();
  hojaInforme.getRange("E4").clearContent(); // Limpiar el QR
}

//acá voy a estar sumando compras, consumos y stock

function actualizarStockConPrecios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaStock = ss.getSheetByName("Stock");
  const hojaCompras = ss.getSheetByName("Compras");
  const hojaConsumos = ss.getSheetByName("Consumos");
  const hojaInsumos = ss.getSheetByName("Insumos");

  // Obtener datos de las hojas
  const datosInsumos = hojaInsumos.getRange(2, 1, hojaInsumos.getLastRow() - 1, 3).getValues();
  const datosCompras = hojaCompras.getRange(2, 1, hojaCompras.getLastRow() - 1, 8).getValues();
  const datosConsumos = hojaConsumos.getRange(2, 1, hojaConsumos.getLastRow() - 1, 5).getValues();

  // Crear un mapa de CDB Insumo para almacenar stock y precios
  let stockPorCDB = {};
  datosInsumos.forEach(fila => {
    const idInsumo = fila[0];
    const nombreInsumo = fila[1];
    const cdbInsumo = fila[2];

    if (cdbInsumo) {
      stockPorCDB[cdbInsumo] = {
        idInsumo,
        nombreInsumo,
        stockActual: 0,
        precioCompra: 0,
        precioVenta: 0
      };
    }
  });

  // Sumar las compras y actualizar precios
  datosCompras.forEach(fila => {
    const cdbInsumo = fila[7]; // CDB Insumo en la columna 8 de Compras
    const cantidadCompra = parseFloat(fila[3]) || 0; // Cantidad en la columna 4 de Compras
    const precioCompra = parseFloat(fila[4]) || 0; // Precio de Compra en la columna 5 de Compras
    const precioVenta = parseFloat(fila[5]) || 0; // Precio de Venta en la columna 6 de Compras

    if (stockPorCDB[cdbInsumo]) {
      // Aumentar stock según la cantidad comprada
      stockPorCDB[cdbInsumo].stockActual += cantidadCompra;

      // Actualizar el precio de compra y venta con el último registrado
      stockPorCDB[cdbInsumo].precioCompra = precioCompra;
      stockPorCDB[cdbInsumo].precioVenta = precioVenta;
    }
  });

  // Restar los consumos del stock
  datosConsumos.forEach(fila => {
    const cdbInsumo = fila[4]; // CDB Insumo en la columna 5 de Consumos
    const cantidadConsumo = parseFloat(fila[3]) || 0; // Cantidad en la columna 4 de Consumos

    if (stockPorCDB[cdbInsumo]) {
      // Restar la cantidad consumida del stock actual
      stockPorCDB[cdbInsumo].stockActual -= cantidadConsumo;
    }
  });

  // Crear las filas para la hoja "Stock" con los valores calculados
  const nuevasFilasStock = Object.entries(stockPorCDB).map(([cdbInsumo, { idInsumo, nombreInsumo, stockActual, precioCompra, precioVenta }]) => [
    idInsumo, nombreInsumo, stockActual, precioCompra, precioVenta, cdbInsumo
  ]);

  // Limpiar y actualizar la hoja "Stock"
  hojaStock.clearContents();
  hojaStock.getRange(1, 1, 1, 6).setValues([["ID Insumo", "Nombre Insumo", "Stock Actual", "Precio de Compra", "Precio de Venta", "CDB Insumo"]]);
  hojaStock.getRange(2, 1, nuevasFilasStock.length, 6).setValues(nuevasFilasStock);

  Logger.log("Stock y precios actualizados correctamente.");
}

//venimos OK y  voy a intentar tener la automatización con mi funcion russellit
//vamos a dejar esta función al final para ir incorporando en la secuencia las nuevas funciones
/**
 * Función RussellIT: Ejecuta en secuencia las funciones:
 * 1. estandarizaciondetextos()
 * 2. crearYRegistrarCarpetasParaVehiculos()
 * 3. gestionarTurnosYEnviarInvitaciones()
 * Aguarda 5 segundos entre cada ejecución.
 */


/**
 * Función principal para ejecutar tareas de notificaciones y envío de correos.
 * Solo se ejecuta en horarios específicos.
 */
function RussellITNotificaciones() {
  try {
    const ahora = new Date();
    const dia = ahora.getDay(); // 0: Domingo, 1: Lunes, ..., 6: Sábado
    const hora = ahora.getHours();

    // Ejecutar de lunes a viernes a las 10:00, 14:00, y 19:00 y sábados a las 13:00
    const esDiaHabil = dia >= 1 && dia <= 5 && (hora === 10 || hora === 14 || hora === 19);
    const esSabado = dia === 6 && hora === 13;

    if (esDiaHabil || esSabado) {
      Logger.log("Iniciando tareas de notificaciones y envío de correos...");

      Logger.log("Actualizando Trabajos Listos Para Entregar - Se enviarán emails...");
      generarInformeDeTrabajosHechos(); // Esta función envía correos solo en estos horarios
      Utilities.sleep(5000);

      Logger.log("Gestionando turnos y enviando invitaciones...");
      gestionarTurnosEnCalendario();

      Logger.log("Ocultando hojas que no deben estar visibles...");
      ocultarTodasLasHojasExceptoModeloPresupuesto();

      Logger.log("Proceso RussellITNotificaciones completado con éxito.");
    } else {
      Logger.log("Fuera del horario de notificaciones y envío de correos. No se ejecuta RussellITNotificaciones.");
    }
  } catch (error) {
    Logger.log("Error en RussellITNotificaciones: " + error.message);
    SpreadsheetApp.getUi().alert("Se produjo un error en RussellITNotificaciones: " + error.message);
  }
}

function RussellITProcesosGenerales() {
  try {
    const ahora = new Date();
    const dia = ahora.getDay();
    const hora = ahora.getHours();

    // Ejecutar solo de lunes a viernes entre las 10:00 y las 19:00, y los sábados a las 10:00 y a las 13:00
    const esDiaHabil = dia >= 1 && dia <= 5 && hora >= 10 && hora <= 21;
    const esSabado = dia === 6 && (hora === 10 || hora === 14);

    if (esDiaHabil || esSabado) {
      Logger.log("Iniciando procesos generales...");

      Logger.log("Estandarizando textos...");
      estandarizaciondetextos();
      Utilities.sleep(5000);

      Logger.log("Estandarizando números de WhatsApp...");
      estandarizarWhatsAppClientes();
      Utilities.sleep(5000);

      Logger.log("Creando y registrando carpetas para vehículos...");
      crearYRegistrarCarpetasParaVehiculos();
      Utilities.sleep(5000);

      Logger.log("Actualizando deudas en trabajos...");
      actualizarDeudaTrabajos();
      Utilities.sleep(5000);

      Logger.log("Actualizando Reporte de Deuda...");
      generarReporteDeuda();
      Utilities.sleep(5000);

      Logger.log("actualizarCarteraDeCheques...");
      actualizarCarteraDeCheques();
      Utilities.sleep(5000);      

      Logger.log("enviarAlertaCheques...");
      enviarAlertaCheques();
      Utilities.sleep(5000);

      Logger.log("Actualizando STOCK - Suma compras y resta la sumatoria de Consumos...");
      actualizarStockConPrecios();

      Logger.log("Proceso RussellITProcesosGenerales completado con éxito.");
    } else {
      Logger.log("Fuera del horario de procesos generales. No se ejecuta RussellITProcesosGenerales.");
    }
  } catch (error) {
    Logger.log("Error en RussellITProcesosGenerales: " + error.message);
    SpreadsheetApp.getUi().alert("Se produjo un error en RussellITProcesosGenerales: " + error.message);
  }
}

//para ejecucón manual - solo la usa Fernando Abeleira - Russell IT
function EjecutarRussellITManual() {
  try {
    Logger.log("Iniciando ejecución manual de procesos RussellIT...");

    // ** Procesos Generales **
    Logger.log("=== Ejecutando RussellITProcesosGenerales ===");

    Logger.log("Estandarizando textos...");
    estandarizaciondetextos();
    Utilities.sleep(3000);

    Logger.log("Estandarizando números de WhatsApp...");
    estandarizarWhatsAppClientes();
    Utilities.sleep(3000);

    Logger.log("Creando y registrando carpetas para vehículos...");
    crearYRegistrarCarpetasParaVehiculos();
    Utilities.sleep(3000);

    Logger.log("Actualizando deudas en trabajos...");
    actualizarDeudaTrabajos();
    Utilities.sleep(3000);

    Logger.log("Generando Reporte de Deuda...");
    generarReporteDeuda();
    Utilities.sleep(3000);

    Logger.log("Actualizando STOCK...");
    actualizarStockConPrecios();
    Utilities.sleep(3000);

    Logger.log("RussellITProcesosGenerales ejecutado con éxito.");

    // ** Notificaciones **
    Logger.log("=== Ejecutando RussellITNotificaciones ===");

    Logger.log("Generando informe de trabajos hechos y enviando correos...");
    generarInformeDeTrabajosHechos();
    Utilities.sleep(3000);

    Logger.log("Gestionando turnos en el calendario...");
    gestionarTurnosEnCalendario();
    Utilities.sleep(3000);

    Logger.log("RussellITNotificaciones ejecutado con éxito.");

    Logger.log("Ejecución manual de procesos RussellIT completada exitosamente.");
  } catch (error) {
    Logger.log("Error en EjecutarRussellITManual: " + error.message);
    SpreadsheetApp.getUi().alert("Se produjo un error en EjecutarRussellITManual: " + error.message);
  }
}


//Acá sigue script definido antes del 16/11/2024

function configurarTriggersRussellIT() {
  eliminarTriggersRussellIT(); // Eliminar triggers existentes antes de configurar nuevos

  // Configuración para RussellITNotificaciones
  for (let day = ScriptApp.WeekDay.MONDAY; day <= ScriptApp.WeekDay.FRIDAY; day++) {
    [10, 14, 19].forEach(function(hour) {
      ScriptApp.newTrigger("RussellITNotificaciones")
        .timeBased()
        .onWeekDay(day)
        .atHour(hour)
        .create();
    });
  }

  // Configuración para RussellITProcesosGenerales
  // De lunes a viernes cada 2 horas entre las 10:00 y las 19:00 (10,12,14,16,18)
  for (let day = ScriptApp.WeekDay.MONDAY; day <= ScriptApp.WeekDay.FRIDAY; day++) {
    for (let hour = 10; hour <= 18; hour += 2) {
      ScriptApp.newTrigger("RussellITProcesosGenerales")
        .timeBased()
        .onWeekDay(day)
        .atHour(hour)
        .create();
    }
  }

  // Sábados cada 2 horas entre las 10:00 y las 14:00 (10,12,14)
  [10, 12, 14].forEach(function(hour) {
    ScriptApp.newTrigger("RussellITProcesosGenerales")
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.SATURDAY)
      .atHour(hour)
      .create();
  });
}

function eliminarTriggersRussellIT() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    const functionName = trigger.getHandlerFunction();
    if (functionName === "RussellITNotificaciones" || functionName === "RussellITProcesosGenerales") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}


function estandarizarWhatsAppClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaClientes = ss.getSheetByName("Base de Clientes");
  const lastRow = hojaClientes.getLastRow();

  // Obtener los datos de la columna H (WhatsApp) sin encabezados
  const rangoWhatsApp = hojaClientes.getRange(2, 8, lastRow - 1, 1); 
  const datosWhatsApp = rangoWhatsApp.getValues();

  // Iterar sobre cada número de WhatsApp
  const datosEstandarizados = datosWhatsApp.map(fila => {
    let numero = fila[0].toString().replace(/\D/g, ''); // Eliminar todos los caracteres que no son números

    // Verificar si el número comienza con "549", si no, añadirlo
    if (!numero.startsWith("549")) {
      numero = "549" + numero;
    }

    return [numero];
  });

  // Actualizar la columna H con los datos estandarizados
  rangoWhatsApp.setValues(datosEstandarizados);
  Logger.log("Números de WhatsApp estandarizados correctamente.");
}


function generarBalanceMensualAnteriorDepuracion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaEgresos = ss.getSheetByName("Egresos Mensuales");
  const hojaTrabajos = ss.getSheetByName("Trabajos Hechos");
  const hojaDatosTaller = ss.getSheetByName("Datos Taller");
  const hojaModelo = ss.getSheetByName("Modelo Balance");

  if (!hojaEgresos || !hojaTrabajos || !hojaDatosTaller || !hojaModelo) {
    throw new Error("Error: Asegúrate de que las hojas 'Egresos Mensuales', 'Trabajos Hechos', 'Datos Taller' y 'Modelo Balance' existan en el archivo.");
  }

  // Obtener datos del taller
  const nombreTaller = hojaDatosTaller.getRange("B2").getValue().trim();
  const fechaActual = new Date();
  const mesAnterior = fechaActual.getMonth() === 0 ? 11 : fechaActual.getMonth() - 1;
  const anioAnterior = fechaActual.getMonth() === 0 ? fechaActual.getFullYear() - 1 : fechaActual.getFullYear();
  const nombreMesAnterior = new Date(anioAnterior, mesAnterior).toLocaleString('es-ES', { month: 'long' });

  // Mostrar la hoja del modelo para trabajar
  hojaModelo.showSheet();
  hojaModelo.getRange("D2").setValue(nombreMesAnterior.charAt(0).toUpperCase() + nombreMesAnterior.slice(1)); // Mes
  hojaModelo.getRange("F2").setValue(anioAnterior); // Año

  // Procesar los datos de egresos y trabajos
  const datosEgresos = hojaEgresos.getRange(2, 1, hojaEgresos.getLastRow() - 1, hojaEgresos.getLastColumn()).getValues();
  const datosTrabajos = hojaTrabajos.getRange(2, 1, hojaTrabajos.getLastRow() - 1, hojaTrabajos.getLastColumn()).getValues();

  let totalEgresos = 0;
  let totalTrabajosLubricentro = 0;
  let totalTrabajosTaller = 0;
  let cantidadEgresos = 0;
  let cantidadLubricentro = 0;
  let cantidadTaller = 0;

  // Filtrar y sumar egresos del mes anterior
  datosEgresos.forEach((fila) => {
    const fechaEgreso = fila[1];
    const montoEgreso = parseFloat(fila[4]) || 0; // Suponiendo que el monto está en la columna 5

    if (fechaEgreso instanceof Date && fechaEgreso.getMonth() === mesAnterior && fechaEgreso.getFullYear() === anioAnterior) {
      totalEgresos += montoEgreso;
      cantidadEgresos++;
    }
  });

  // Filtrar y sumar trabajos del mes anterior
  datosTrabajos.forEach((fila) => {
    const fechaTrabajo = fila[6];
    const sector = fila[3];
    const montoTrabajo = parseFloat(fila[9]) || 0;

    if (fechaTrabajo instanceof Date && fechaTrabajo.getMonth() === mesAnterior && fechaTrabajo.getFullYear() === anioAnterior) {
      if (sector === "Lubricentro") {
        totalTrabajosLubricentro += montoTrabajo;
        cantidadLubricentro++;
      } else if (sector === "Taller") {
        totalTrabajosTaller += montoTrabajo;
        cantidadTaller++;
      }
    }
  });

  // Calcular ingresos totales y balance
  const totalIngresos = totalTrabajosLubricentro + totalTrabajosTaller;
  const balanceFinal = totalIngresos - totalEgresos;

  // Escribir resultados en la hoja 'Modelo Balance'
  hojaModelo.getRange("D3").setValue(cantidadEgresos); // Cantidad de egresos
  hojaModelo.getRange("D4").setValue(formatCurrency(totalEgresos)); // Total egresos
  hojaModelo.getRange("D6").setValue(cantidadLubricentro); // Cantidad trabajos Lubricentro
  hojaModelo.getRange("D7").setValue(formatCurrency(totalTrabajosLubricentro)); // Total trabajos Lubricentro
  hojaModelo.getRange("D9").setValue(cantidadTaller); // Cantidad trabajos Taller
  hojaModelo.getRange("D10").setValue(formatCurrency(totalTrabajosTaller)); // Total trabajos Taller
  hojaModelo.getRange("D12").setValue(formatCurrency(totalEgresos)); // Total egresos (repetido para claridad)
  hojaModelo.getRange("D13").setValue(formatCurrency(totalIngresos)); // Total ingresos
  hojaModelo.getRange("D14").setValue(formatCurrency(balanceFinal)); // Balance final

  Logger.log(`Egresos totales para ${nombreMesAnterior} ${anioAnterior}: ${totalEgresos}`);
  Logger.log(`Ingresos totales para ${nombreMesAnterior} ${anioAnterior}: ${totalIngresos}`);
  Logger.log(`Balance final para ${nombreMesAnterior} ${anioAnterior}: ${balanceFinal}`);

  // Generar PDF del balance
  const nombrePDF = `${nombreTaller}_Balance_${nombreMesAnterior}_${anioAnterior}.pdf`;
  const pdfBlob = crearPDFDesdeHoja(hojaModelo, nombrePDF);

  // Guardar el PDF en la carpeta "Balances"
  const carpetaPrincipal = DriveApp.getFolderById(ss.getId()).getParents().next();
  let carpetaBalances = carpetaPrincipal.getFoldersByName("Balances");
  if (!carpetaBalances.hasNext()) {
    carpetaBalances = carpetaPrincipal.createFolder("Balances");
  } else {
    carpetaBalances = carpetaBalances.next();
  }
  carpetaBalances.createFile(pdfBlob);

  Logger.log(`PDF generado y guardado como '${nombrePDF}' en la carpeta 'Balances'.`);

  // Ocultar la hoja del modelo
  hojaModelo.hideSheet();
}

// Función para formatear números como moneda
function formatCurrency(value) {
  return `\$${Math.round(value).toLocaleString("es-AR")}`;
}

// Función para crear un PDF desde una hoja
function crearPDFDesdeHoja(hoja, nombreArchivo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaId = hoja.getSheetId();
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${hojaId}`;

  const response = UrlFetchApp.fetch(url, {
    method: "GET",
    headers: {
      "Authorization": `Bearer ${ScriptApp.getOAuthToken()}`
    }
  });

  const blob = response.getBlob().setName(nombreArchivo);
  return blob;
}

function enviarMailBalanceMensual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDatosTaller = ss.getSheetByName("Datos Taller");

  if (!hojaDatosTaller) {
    throw new Error("Error: Asegúrate de que la hoja 'Datos Taller' exista en el archivo.");
  }

  const nombreTaller = "DGS"; // Usar el nombre fijo "DGS" para el taller
  const emailTaller = hojaDatosTaller.getRange("B5").getValue().trim();

  if (!emailTaller) {
    throw new Error("Error: Asegúrate de que el email esté completo en la hoja 'Datos Taller'.");
  }

  // Calcular el mes y año anterior
  const fechaActual = new Date();
  const mesAnterior = fechaActual.getMonth() === 0 ? 11 : fechaActual.getMonth() - 1;
  const anioAnterior = fechaActual.getMonth() === 0 ? fechaActual.getFullYear() - 1 : fechaActual.getFullYear();
  const nombreMesAnterior = new Date(anioAnterior, mesAnterior).toLocaleString('es-ES', { month: 'long' });

  // Formatear el nombre del archivo como "DGS_Balance_<mes>_<año>.pdf"
  const nombreArchivoBuscar = `DGS_Balance_${nombreMesAnterior}_${anioAnterior}.pdf`;

  // Buscar la carpeta "Balances"
  const carpetaPrincipal = DriveApp.getFolderById(ss.getId()).getParents().next();
  const carpetaBalances = carpetaPrincipal.getFoldersByName("Balances").next();

  // Buscar el archivo del mes anterior
  const archivos = carpetaBalances.getFilesByName(nombreArchivoBuscar);
  let archivoPDF = null;

  while (archivos.hasNext()) {
    archivoPDF = archivos.next();
    break;
  }

  if (!archivoPDF) {
    Logger.log(`No se encontró un archivo para el balance de ${nombreMesAnterior} ${anioAnterior}.`);
    return;
  }

  // Configurar asunto y cuerpo del correo
  const asunto = `Balance del mes de ${nombreMesAnterior} ${anioAnterior}`;
  const cuerpo = `
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
      <h2 style="color: #0056b3;">Estimado equipo de ${nombreTaller},</h2>
      <p>
        Por medio del presente, les compartimos el balance correspondiente al mes de <strong>${nombreMesAnterior} ${anioAnterior}</strong>.
        Podrán encontrarlo adjunto en este correo para su revisión.
      </p>
      <p>
        Si tienen alguna consulta o requieren más información, no duden en contactarnos. Estamos a su disposición.
      </p>
      <p style="margin-top: 30px;">Saludos cordiales,</p>
      <p>
        <strong>Sistema Mecamanagement by Russell IT</strong><br>
        <em>Automatización y soporte para su taller</em>
      </p>
    </div>
  `;

  // Enviar el correo con el archivo adjunto
  MailApp.sendEmail({
    to: emailTaller,
    subject: asunto,
    htmlBody: cuerpo,
    attachments: [archivoPDF.getAs(MimeType.PDF)],
    name: "Sistema Mecamanagement by Russell IT"
  });

  Logger.log(`Correo enviado correctamente con el archivo '${nombreArchivoBuscar}' adjunto.`);
}


function NuevoTaller() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreArchivo = ss.getName();
  let log = []; // Registro de acciones realizadas para el proceso

  // Paso 1: Crear la carpeta de Vehículos
  const carpetaPrincipal = DriveApp.getFileById(ss.getId()).getParents().next();
  const carpetaVehiculos = carpetaPrincipal.createFolder("Vehículos");
  const idCarpetaVehiculos = carpetaVehiculos.getId();

  SpreadsheetApp.getUi().alert(`Carpeta 'Vehículos' creada. ID: ${idCarpetaVehiculos}. 
    Debes reemplazar este ID en el script.`);

  // Paso 2: Instrucción para crear un calendario y solicitar el ID
  const idCalendar = SpreadsheetApp.getUi().prompt("Crear Calendario",
    "Por favor, crea un nuevo calendario en Google Calendar y coloca aquí el ID:", SpreadsheetApp.getUi().ButtonSet.OK).getResponseText();

  if (!idCalendar) {
    SpreadsheetApp.getUi().alert("Debe ingresar un ID de calendario para continuar.");
    return;
  }
  
  log.push(`ID de Calendario ingresado: ${idCalendar}`);

  // Paso 3: Limpiar las hojas de datos, excepto las de templates
  const hojasParaLimpiar = [
    { nombre: "Base de Clientes", desdeFila: 2 },
    { nombre: "Base Vehículos", desdeFila: 2 },
    { nombre: "Trabajos Hechos", desdeFila: 2 },
    { nombre: "Turnos", desdeFila: 2 },
    { nombre: "Presupuestos", desdeFila: 2 },
    { nombre: "Carpetas", desdeFila: 2 },
    { nombre: "Reporte Deuda", desdeFila: 2 },
    { nombre: "Stock", desdeFila: 3 },
    { nombre: "Compras", desdeFila: 3 },
    { nombre: "Proveedore", desdeFila: 2 },
    { nombre: "Insumos", desdeFila: 3 },
    { nombre: "Consumos", desdeFila: 3 }
  ];

  hojasParaLimpiar.forEach(hojaInfo => {
    const hoja = ss.getSheetByName(hojaInfo.nombre);
    if (hoja) {
      const lastColumn = hoja.getLastColumn();
      const numRows = hoja.getLastRow() - hojaInfo.desdeFila + 1;
      if (numRows > 0) hoja.getRange(hojaInfo.desdeFila, 1, numRows, lastColumn).clearContent();
      log.push(`Hoja ${hojaInfo.nombre} limpiada.`);
    }
  });

  // Paso 4: Completar los datos del taller en la hoja "Datos Taller"
  const hojaDatosTaller = ss.getSheetByName("Datos Taller");
  if (!hojaDatosTaller) {
    SpreadsheetApp.getUi().alert("La hoja 'Datos Taller' no se encontró. Debes crearla para continuar.");
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const nombreTaller = ui.prompt("Configurar Taller", "Ingrese el nombre del taller:", ui.ButtonSet.OK).getResponseText();
  const direccion = ui.prompt("Configurar Taller", "Ingrese la dirección del taller:", ui.ButtonSet.OK).getResponseText();
  const telefono = ui.prompt("Configurar Taller", "Ingrese el teléfono del taller:", ui.ButtonSet.OK).getResponseText();
  const email = ui.prompt("Configurar Taller", "Ingrese el email del taller:", ui.ButtonSet.OK).getResponseText();
  const whatsapp = ui.prompt("Configurar Taller", "Ingrese el número de WhatsApp del taller:", ui.ButtonSet.OK).getResponseText();
  const nombreGerente = ui.prompt("Configurar Taller", "Ingrese el nombre del gerente del taller:", ui.ButtonSet.OK).getResponseText();

  // Ingresar los valores en la hoja "Datos Taller"
  hojaDatosTaller.getRange("B2").setValue(nombreTaller);
  hojaDatosTaller.getRange("B3").setValue(direccion);
  hojaDatosTaller.getRange("B4").setValue(telefono);
  hojaDatosTaller.getRange("B5").setValue(email);
  hojaDatosTaller.getRange("B6").setValue(whatsapp);
  hojaDatosTaller.getRange("B7").setValue(nombreGerente);

  log.push("Datos del taller configurados.");

  // Paso 5: Ocultar todas las hojas excepto "Modelo Presupuesto"
  const hojas = ss.getSheets();
  hojas.forEach(hoja => {
    if (hoja.getName() !== "Modelo Presupuesto") hoja.hideSheet();
  });

  // Paso 6: Recordatorio final para reemplazar IDs
  SpreadsheetApp.getUi().alert(
    "Configuración inicial del nuevo taller completada.\n\nNanu acordate reemplazar los ID de carpeta Vehículos y del calendario en el script, con las instrucciones que aparecen en el código.\n\nTenes que crear el calendario.\n\nTenes que Campartir carpeta con tu cliente.\n\nTenes que configurar los triggers de notificaciones diario cada 4 horas y el de procesos generales cada 2 horas.\n\nNo tenes que preocuparte por la repetición de la funciones estan limitadas a dias y horarios hábiles.\n\nTe queda la configuracion personalizada de los reportes. \n\nCompartir accesos mobile y desktop para que el cliente trabaje.\n\nFelicitaciones Nanu - vamos por mas ventas."
  );

  // Log de finalización
  Logger.log("Proceso 'Nuevo Taller' completado con éxito.");
  log.forEach(accion => Logger.log(accion));
}

function mostrarTodasLasHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = ss.getSheets();
  
  hojas.forEach(hoja => {
    if (hoja.isSheetHidden()) {
      hoja.showSheet();
      Logger.log(`Hoja mostrada: ${hoja.getName()}`);
    } else {
      Logger.log(`Hoja ya visible: ${hoja.getName()}`);
    }
  });

  Logger.log("Todas las hojas están visibles.");
}

function ocultarTodasLasHojasExceptoModeloPresupuesto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = ss.getSheets();

  hojas.forEach(hoja => {
    // Mantener visibles las hojas especificadas
    if (hoja.getName() !== "Modelo Presupuesto" && 
        hoja.getName() !== "Reporte Pendientes" && 
        hoja.getName() !== "Orden de Reparación" && 
        hoja.getName() !== "Condiciones Generales" && 
        hoja.getName() !== "Presupuesto Lubricentro") {
      hoja.hideSheet();
      Logger.log(`Hoja oculta: ${hoja.getName()}`);
    } else {
      Logger.log(`Hoja mantenida visible: ${hoja.getName()}`);
    }
  });

  Logger.log("Todas las hojas excepto 'Modelo Presupuesto', 'Reporte Pendientes' y 'Presupuesto Lubricentro' han sido ocultadas.");
}


function generarReportePendientes() {
  const hojaDatos = "Trabajos Hechos";
  const hojaReporte = "Reporte Pendientes";
  const celdaCliente = "B12";
  const filaInicioReporte = 19;
  const filaFinReporte = 27;
  const celdaSuma = "G28";

  // ID de la carpeta en Google Drive para guardar los PDFs
  const folderId = "1N-gMRjh8jUqTzVsBTPPmWXEB5h5CkvvQ";

  // Obtener las hojas
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const datosSheet = ss.getSheetByName(hojaDatos);
  const reporteSheet = ss.getSheetByName(hojaReporte);

  // Obtener el nombre del cliente
  const cliente = reporteSheet.getRange(celdaCliente).getValue();

  if (!cliente) {
    SpreadsheetApp.getUi().alert("Por favor, ingresa un cliente en la celda B12.");
    return;
  }

  // Limpiar las celdas B19:G27 y la celda G28
  reporteSheet.getRange(filaInicioReporte, 2, filaFinReporte - filaInicioReporte + 1, 6).clearContent();
  reporteSheet.getRange(celdaSuma).clearContent();

  // Leer los datos de "Trabajos Hechos"
  const datos = datosSheet.getDataRange().getValues();
  let filaReporte = filaInicioReporte;

  // Filtrar y copiar los datos al reporte
  datos.forEach((fila, index) => {
    if (index === 0) return; // Saltar la fila de encabezados
    const [id, patente, nombreCliente, sector, descripcion, fechaInicio, fechaFin, estado, debe, monto, pago] = fila;

    // Verificar condiciones: cliente coincide, debe > 0, estado es "04_ENTREGADO"
    if (nombreCliente === cliente && debe > 0 && estado === "04_ENTREGADO") {
      const fechaFinFormateada = Utilities.formatDate(new Date(fechaFin), Session.getScriptTimeZone(), "dd/MM/yyyy");
      reporteSheet.getRange(filaReporte, 2).setValue(descripcion + " - " + fechaFinFormateada); // Columna B
      reporteSheet.getRange(filaReporte, 5).setValue(monto); // Columna E
      reporteSheet.getRange(filaReporte, 6).setValue(pago); // Columna F
      reporteSheet.getRange(filaReporte, 7).setValue(debe); // Columna G
      filaReporte++;
    }
  });

  if (filaReporte === filaInicioReporte) {
    SpreadsheetApp.getUi().alert(`No se encontraron trabajos adeudados para el cliente: ${cliente}`);
    return;
  }

  // Calcular y escribir la sumatoria en la celda G28 antes de generar el PDF
  const rangoDeuda = reporteSheet.getRange(filaInicioReporte, 7, filaFinReporte - filaInicioReporte + 1);
  const sumaDeuda = rangoDeuda.getValues().flat().reduce((total, valor) => total + (parseFloat(valor) || 0), 0);
  reporteSheet.getRange(celdaSuma).setValue(sumaDeuda);

  // Asegurar que el cálculo esté completo antes de continuar
  SpreadsheetApp.flush();

  // Esperar 2 segundos antes de generar el PDF
  Utilities.sleep(2000);

  // Generar el nombre del archivo PDF
  const fechaActual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy");
  const nombrePDF = `Reporte_${cliente}_${fechaActual}.pdf`;

  // Guardar como PDF en Google Drive
  const carpeta = DriveApp.getFolderById(folderId);
  const pdfBlob = convertirHojaAPDF(ss, hojaReporte);
  const archivo = carpeta.createFile(pdfBlob.setName(nombrePDF));

  SpreadsheetApp.getUi().alert(`El archivo ha sido guardado correctamente en la carpeta "Reporte Deudas".`);
}

// Función auxiliar para convertir una hoja a PDF
function convertirHojaAPDF(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`La hoja '${sheetName}' no existe.`);
  }

  const sheetId = sheet.getSheetId();
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=pdf&gid=${sheetId}&size=letter&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`;

  const token = ScriptApp.getOAuthToken();

  const options = {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() !== 200) {
    const errorDetails = response.getContentText();
    throw new Error(`Error al generar el PDF. Código de respuesta: ${response.getResponseCode()}. Detalles: ${errorDetails}`);
  }

  return response.getBlob();
}


// ahora empezamos con temas del lubricentro

function envioYRegistroDePresupuestoLubricentro() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaPresupuestoLubricentro = ss.getSheetByName("Presupuesto Lubricentro");
    const hojaPresupuestos = ss.getSheetByName("Presupuestos");
    const hojaDatosTaller = ss.getSheetByName("Datos Taller");

    // Obtener datos del presupuesto
    const cliente = hojaPresupuestoLubricentro.getRange("B12").getValue();
    const emailCliente = hojaPresupuestoLubricentro.getRange("B13").getValue();
    const patente = hojaPresupuestoLubricentro.getRange("D12").getValue();
    const monto = hojaPresupuestoLubricentro.getRange("G28").getValue() || 0;
    const fechaEnvio = hojaPresupuestoLubricentro.getRange("B9").getValue();
    const fechaVencimiento = hojaPresupuestoLubricentro.getRange("F15").getValue();
    const nombrePDF = "LUBRI_" + hojaPresupuestoLubricentro.getRange("F12").getValue();

    // Obtener datos del taller
    const nombreTaller = hojaDatosTaller.getRange("B2").getValue();
    const direccionTaller = hojaDatosTaller.getRange("B3").getValue();
    const telefonoTaller = hojaDatosTaller.getRange("B4").getValue();
    const emailTaller = hojaDatosTaller.getRange("B5").getValue();
    const whatsappTaller = hojaDatosTaller.getRange("B6").getValue();
    const nombreGerente = hojaDatosTaller.getRange("B7").getValue();

    // Confirmar envío
    const ui = SpreadsheetApp.getUi();
    const respuesta = ui.alert(
      'Confirmación',
      `Está a punto de enviar un presupuesto a "${cliente}". ¿Está seguro?`,
      ui.ButtonSet.YES_NO
    );

    if (respuesta !== ui.Button.YES) {
      Logger.log("Acción cancelada por el usuario.");
      return;
    }

    // Generar el código QR usando QuickChart
    const datoParaQR = patente;
    if (!datoParaQR) {
      throw new Error("El dato en la celda D12 (Patente) está vacío. No se puede generar el QR.");
    }

    const urlQR = `https://quickchart.io/qr?text=${encodeURIComponent(datoParaQR)}&size=150`;
    const imagenQR = UrlFetchApp.fetch(urlQR).getBlob();
    hojaPresupuestoLubricentro.getRange("E4").setValue(""); // Limpiar cualquier valor anterior
    hojaPresupuestoLubricentro.insertImage(imagenQR, 5, 4).setAnchorCell(hojaPresupuestoLubricentro.getRange("E4"));

    // Generar carpeta de presupuestos si no existe
    const carpetaPrincipalId = "1I3BoZoPdCwwLcDRbxi-Ez6c4XX0txsPR";
    let carpetaPresupuestos = DriveApp.getFolderById(carpetaPrincipalId).getFoldersByName("Presupuestos Lubricentro");

    if (!carpetaPresupuestos.hasNext()) {
      carpetaPresupuestos = DriveApp.getFolderById(carpetaPrincipalId).createFolder("Presupuestos Lubricentro");
    } else {
      carpetaPresupuestos = carpetaPresupuestos.next();
    }

    // Generar el PDF
    const pdfBlob = generarPDF(hojaPresupuestoLubricentro, nombrePDF);
    const pdfFile = carpetaPresupuestos.createFile(pdfBlob).setName(`${nombrePDF}.pdf`);
    const linkPDF = pdfFile.getUrl();

    // Registrar el presupuesto
    const ultimaFila = hojaPresupuestos.getLastRow();
    const nuevoIdPresupuesto = "LUBRI_" + Utilities.formatString("%04d", ultimaFila);

    hojaPresupuestos.appendRow([
      nuevoIdPresupuesto, patente || "0", cliente, linkPDF, monto, fechaEnvio, "PRESUPUESTADO"
    ]);

    // Preparar y enviar el correo electrónico
    const asunto = `Presupuesto - ${nombrePDF}`;
    const cuerpo = `
      Estimad@ ${cliente},<br><br>
      Adjuntamos el presupuesto solicitado.<br><br>
      Tenga en cuenta que este presupuesto tiene validez por 7 días, venciendo el día ${fechaVencimiento}.<br><br>
      Cualquier inquietud no dude en contactarnos:<br>
      Teléfono: <a href="tel:${telefonoTaller}">${telefonoTaller}</a><br>
      WhatsApp: <a href="https://wa.me/549${whatsappTaller}">${whatsappTaller}</a><br>
      Email: <a href="mailto:${emailTaller}">${emailTaller}</a><br><br>
      Dirección: ${direccionTaller}<br><br>
      Saludos,<br><br>
      ${nombreGerente}<br>
      ${nombreTaller}
    `;

    if (emailCliente) {
      MailApp.sendEmail({
        to: emailCliente,
        subject: asunto,
        htmlBody: cuerpo,
        attachments: [pdfFile.getAs(MimeType.PDF)],
        name: nombreTaller
      });
      ui.alert(`Se envió el presupuesto al mail "${emailCliente}"`);
    } else {
      ui.alert("No se encontró un correo electrónico del cliente. Presupuesto registrado en la carpeta 'Presupuestos Lubricentro'.");
    }

    Logger.log("Presupuesto enviado o registrado exitosamente.");
  } catch (error) {
    Logger.log("Error en envioYRegistroDePresupuestoLubricentro: " + error.message);
    SpreadsheetApp.getUi().alert("Se produjo un error: " + error.message);
  }
}

// Función para generar el PDF
function generarPDF(hoja, nombrePDF) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaId = hoja.getSheetId();
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${hojaId}`;

  const response = UrlFetchApp.fetch(url, {
    method: "GET",
    headers: {
      "Authorization": `Bearer ${ScriptApp.getOAuthToken()}`
    }
  });

  return response.getBlob().setName(`${nombrePDF}.pdf`);
}

//arrancamos con la cartera de cheques - hasta la fila 1847 tenemos todo funcionando maravillosamente

function actualizarCarteraDeCheques() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTrabajos = ss.getSheetByName("Trabajos Hechos");
  const hojaCartera = ss.getSheetByName("Cartera de Cheques");

  // Obtener todos los datos de la hoja "Trabajos Hechos"
  const ultimaFilaTrabajos = hojaTrabajos.getLastRow();
  const ultimaColumnaTrabajos = hojaTrabajos.getLastColumn();
  if (ultimaFilaTrabajos < 2) {
    Logger.log("La hoja 'Trabajos Hechos' no tiene datos suficientes.");
    return;
  }

  const datosTrabajos = hojaTrabajos.getRange(2, 1, ultimaFilaTrabajos - 1, ultimaColumnaTrabajos).getValues();

  // Filtrar solo filas que tienen datos válidos
  const filasConDatos = datosTrabajos.filter(fila => fila.some(celda => celda !== ""));

  if (filasConDatos.length === 0) {
    Logger.log("No hay filas con datos en 'Trabajos Hechos'.");
    return;
  }

  // Obtener datos actuales de "Cartera de Cheques"
  const ultimaFilaCartera = hojaCartera.getLastRow();
  let datosCartera = [];
  if (ultimaFilaCartera > 1) {
    datosCartera = hojaCartera.getRange(2, 1, ultimaFilaCartera - 1, hojaCartera.getLastColumn()).getValues();
  }

  const chequesExistentes = datosCartera.map(fila => fila[1]); // Nº Cheque en la columna 2

  // Filtrar nuevos cheques para agregar
  const nuevosCheques = filasConDatos
    .filter(fila => fila[15] && !chequesExistentes.includes(fila[15])) // Columna 16: Nº Cheque
    .map(fila => [
      `CHQ-${Date.now()}`, // IDCheque único
      fila[15], // Nº Cheque
      fila[16], // Nombre Cheque
      fila[17], // Fecha Cobro Cheque
      fila[18], // Importe Cheque
      fila[0],  // ID Trabajo
      fila[1],  // Patente
      fila[2],  // Cliente
      fila[12], // Informe
      "Pendiente" // Estado del correo
    ]);

  // Verificar si hay nuevos cheques para agregar
  if (nuevosCheques.length > 0) {
    hojaCartera.getRange(ultimaFilaCartera + 1, 1, nuevosCheques.length, nuevosCheques[0].length).setValues(nuevosCheques);
    Logger.log(`${nuevosCheques.length} cheques agregados a la cartera.`);
  } else {
    Logger.log("No hay cheques nuevos para agregar.");
  }
}



function enviarAlertaCheques() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaCartera = ss.getSheetByName("Cartera de Cheques");
  const hojaDatosTaller = ss.getSheetByName("Datos Taller"); // Hoja donde se encuentra el email del taller

  // Obtener el email del taller desde la celda correspondiente
  const emailTaller = hojaDatosTaller.getRange("B5").getValue().trim();
  if (!validarEmail(emailTaller)) {
    Logger.log(`El email "${emailTaller}" no es válido. Revisa la hoja 'Datos Taller'.`);
    throw new Error("El email del taller no es válido.");
  }

  // Obtener datos de la hoja "Cartera de Cheques"
  const hoy = new Date();
  const ultimaFilaCartera = hojaCartera.getLastRow();
  if (ultimaFilaCartera < 2) {
    Logger.log("La hoja 'Cartera de Cheques' no tiene datos suficientes.");
    return;
  }

  const datosCartera = hojaCartera.getRange(2, 1, ultimaFilaCartera - 1, hojaCartera.getLastColumn()).getValues();
  const chequesHoy = datosCartera.filter(fila => {
    const fechaCobro = fila[3]; // Columna 4: Fecha Cobro Cheque
    const estadoCorreo = fila[9]; // Columna 10: Estado del correo
    return fechaCobro instanceof Date && fechaCobro.toDateString() === hoy.toDateString() && estadoCorreo === "Pendiente";
  });

  if (chequesHoy.length > 0) {
    const asunto = "🔔 Alerta: Cheques por cobrar hoy";

    // Construir el cuerpo del correo en HTML
    let mensajeHTML = `
      <div style="font-family: Arial, sans-serif; line-height: 1.5; color: #333;">
        <h2 style="color: #4CAF50; font-size: 20px; margin-bottom: 10px;">Reporte de Cheques por Cobrar</h2>
        <p>Estimado equipo de <strong>${emailTaller}</strong>,</p>
        <p>Se detalla a continuación el listado de cheques que tienen fecha de cobro para el día de hoy:</p>
        <table style="width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 14px;">
          <thead>
            <tr style="background-color: #f2f2f2;">
              <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Nº Cheque</th>
              <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Nombre Cheque</th>
              <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Fecha Cobro</th>
              <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Importe</th>
              <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Cliente</th>
              <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Patente</th>
            </tr>
          </thead>
          <tbody>`;

    chequesHoy.forEach(fila => {
      mensajeHTML += `
            <tr>
              <td style="border: 1px solid #ddd; padding: 8px;">${fila[1]}</td>
              <td style="border: 1px solid #ddd; padding: 8px;">${fila[2]}</td>
              <td style="border: 1px solid #ddd; padding: 8px;">${fila[3].toLocaleDateString()}</td>
              <td style="border: 1px solid #ddd; padding: 8px;">$${fila[4].toFixed(2)}</td>
              <td style="border: 1px solid #ddd; padding: 8px;">${fila[7]}</td>
              <td style="border: 1px solid #ddd; padding: 8px;">${fila[6]}</td>
            </tr>`;
    });

    mensajeHTML += `
          </tbody>
        </table>
        <p style="margin-top: 20px;">Por favor, asegúrense de realizar el cobro correspondiente a tiempo.</p>
        <p style="margin-top: 30px;">Saludos cordiales,</p>
        <p style="font-weight: bold; color: #333;">Sistema de Gestión Mecamanagement</p>
        <p style="font-size: 12px; color: #888;">(Mensaje generado automáticamente, no responder a este correo)</p>
      </div>`;

    // Configurar el remitente enmascarado
    const opciones = {
      name: "Gestión de Cheques - DGS",
      htmlBody: mensajeHTML
    };

    // Enviar el correo
    MailApp.sendEmail(emailTaller, asunto, "", opciones);
    Logger.log(`Correo enviado a ${emailTaller} con cheques por cobrar hoy.`);

    // Actualizar estado del correo
    chequesHoy.forEach(fila => {
      const index = datosCartera.findIndex(c => c[1] === fila[1]); // Comparar por Nº Cheque
      if (index > -1) {
        hojaCartera.getRange(index + 2, 10).setValue("Enviado"); // Columna 10: Estado del correo
      } else {
        Logger.log(`No se pudo encontrar la fila para el cheque Nº ${fila[1]}`);
      }
    });
  } else {
    Logger.log("No hay cheques por cobrar hoy o ya se enviaron los correos.");
  }
}

/**
 * Validar si un email tiene un formato válido
 * @param {string} email - Dirección de correo electrónico a validar
 * @returns {boolean} - True si el email es válido, false en caso contrario
 */
function validarEmail(email) {
  const patronEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return patronEmail.test(email);
}
