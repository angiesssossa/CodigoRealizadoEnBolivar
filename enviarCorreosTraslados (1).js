function enviarCorreosTraslados() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. Cuadro de diálogo para ingresar datos personales que iran en la firma del correo
  const respuesta = ui.prompt(
    '✏️ Datos para la firma',
    'Ingresa tu nombre y cargo (ej: "Nombre Completo | Perfil a Cargo"):',
    ui.ButtonSet.OK_CANCEL
  );
  
  //al seleccionar cancelar, se detiene el código y no hara cambios
  if (respuesta.getSelectedButton() === ui.Button.CANCEL) return;
  const datosUsuario = respuesta.getResponseText();

  // URL publica del logo de bolivar
  const urlLogo = "https://cdn.brandfetch.io/idNYmlVYed/w/400/h/400/theme/dark/icon.jpeg?c=1dxbfHSJFAPEGdCLU4o5B";

  // Plantilla de firma con imagen a la izquierda y texto a la derecha
  const firmaHTML = `
    <div style="font-family: Arial, sans-serif; color: #333; margin-top: 20px; border-top: 1px solid #eee; padding-top: 15px;">
      <table style="width: 100%; border-collapse: collapse;">
        <tr>
          <td style="width: 80px; vertical-align: top; padding-right: 15px;">
            <img src="${urlLogo}" 
                 alt="Logo Seguros Bolívar" 
                 style="width: 150px; height: auto; display: block;">
          </td>
          <td style="vertical-align: top; padding-left: 15px; border-left: 1px solid #eee;">
            <p style="margin: 10px 0 5px 0; font-weight: bold;">${datosUsuario}</p>
            <p style="margin: 0 0 3px 0; font-size: 11px; line-height: 1.4;">
              <stronger>Dirección Nacional Administrativa ARL</stronger><br>
              Av. El Dorado #68 B-65 Piso 7<br>
              Bogotá, Colombia
            </p>
            <p style="margin: 5px 0 0 0; font-size: 11px;">
              <a href="https://www.segurosbolivar.com" 
                 style="color: #0066cc; text-decoration: none;">www.segurosbolivar.com</a>
            </p>
          </td>
        </tr>
      </table>
    </div>
  `;

  // Busca la hoja llamada Traslados Correos
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TRASLADOS CORREOS");
  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0];

  // Índices de columnas
  const iNit = encabezados.indexOf("NIT");
  const iEmpresa = encabezados.indexOf("EMPRESA");
  const iPoliza = encabezados.indexOf("POLIZA");
  const iNumCaso = encabezados.indexOf("NUMERO CASO");
  const iFechaRadicado = encabezados.indexOf("FECHA RADICADO");
  const iFechaSolicitada = encabezados.indexOf("FECHA SOLICITADA");
  const iLocalidad = encabezados.indexOf("LOCALIDAD");
  const iCanal = encabezados.indexOf("CANAL");
  const iGrupo = encabezados.indexOf("GRUPO");
  const iCanalComercial = encabezados.indexOf("CANAL COMERCIAL");
  const iAsesor = encabezados.indexOf("ASESOR");
  const iPygGlobal = encabezados.indexOf("P&G GLOBAL");
  const iUltAporte = encabezados.indexOf("ULTIMO APORTE");
  const iCorreoEnv = encabezados.indexOf("CORREO ENVIADO");

  //guarda en una lista los correos de destinatarios
  const destinatariosPorGrupo = {
    "Agencias & Grupos Homogeneos": ["angie.sosa@segurosbolivar.com"],
    "Corredores & Agencias Multiples": ["lina.omaira.gonzalez@segurosbolivar.com"]
  };

  //Verifica el envio de correo de una empresa con la condicion "SI"
for (let i = 1; i < datos.length; i++) {
  const fila = datos[i];
  const correoYaEnviado = fila[iCorreoEnv].toString().toUpperCase() === "SI";
  const nit = fila[iNit];

  // Verifica que el correo no se haya enviado y que haya NIT Para el envio
  if (!correoYaEnviado && nit && nit.toString().trim() !== "") {
    const empresa = fila[iEmpresa];
    const poliza = fila[iPoliza];
    const numCaso = fila[iNumCaso];
    const fechaRad = formatearFecha(fila[iFechaRadicado]);
    const fechaSoli = formatearFecha(fila[iFechaSolicitada]);
    const localidad = fila[iLocalidad];
    const canal = fila[iCanal];
    const grupo = fila[iGrupo];
    const canalComercial = fila[iCanalComercial];
    const asesor = fila[iAsesor];
    const pygGlobal = formatearPorcentaje(fila[iPygGlobal]);
    const ultAporte = formatearMoneda(fila[iUltAporte]);

    const asunto = `Traslado ARL Póliza ${poliza} - NIT ${nit} - ${empresa} - CASO ${numCaso}`;
    const mensajeHTML = `Buenos días, <br><br>
Nos permitimos informar que el día ${fechaRad} fue radicada la solicitud de traslado de la empresa <b>${empresa}</b> con NIT <b>${nit}</b> y Póliza <b>${poliza}</b> para surtir efecto a partir del <b>${fechaSoli}</b>.<br><br>
Localidad: ${localidad}<br>
Canal: ${canal}<br>
Canal Comercial: ${canalComercial}<br>
Asesor A y C: ${asesor}<br>
P&G Global: ${pygGlobal}<br>
Último Aporte: ${ultAporte}<br><br><br>
Cordialmente,<br><br>
${firmaHTML}`;

    MailApp.sendEmail({
      to: destinatariosPorGrupo[grupo].join(","),
      subject: asunto,
      htmlBody: mensajeHTML
    });

    hoja.getRange(i + 1, iCorreoEnv + 1).setValue("SI");
  }
}
  
  ui.alert("✅ Correos enviados", "Los correos se han enviado correctamente con tu firma profesional.", ui.ButtonSet.OK);
}

//Funcion para que transcriba fecha a DD/MM/YYYY
function formatearFecha(valor) {
  const fecha = new Date(valor);
  if (!isNaN(fecha)) {
    const utcOffset = fecha.getTimezoneOffset() * 60000;
    const fechaLocal = new Date(fecha.getTime() + utcOffset);
    return Utilities.formatDate(fechaLocal, "America/Bogota", "dd/MM/yyyy");
  }
  return valor;
}

//Funcion para poner el valor en porcentaje
function formatearPorcentaje(valor) {
  const porcentaje = parseFloat(valor);
  if (!isNaN(porcentaje)) {
    return (porcentaje * 100).toFixed(2).replace('.', ',') + '%';
  }
  return valor;
}

//Funcion para poner valor en Pesos
function formatearMoneda(valor) {
  const numero = parseFloat(valor);
  if (!isNaN(numero)) {
    return '$' + numero.toLocaleString('es-CO');
  }
  return valor;
}