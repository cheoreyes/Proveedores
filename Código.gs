/**
 * Servidor: Código.gs
 * Sistema de Control de Proveedores - Repuestos VZLA
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Repuestos VZLA - Control de Proveedores')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * OBTIENE Y NORMALIZA LOS DATOS DIRECTAMENTE DESDE LA PESTAÑA 'DATOS' 
 */
function obtenerDatosProveedoresBackend() {
  try {
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hojaDatos = libro.getSheetByName("DATOS");
    if (!hojaDatos) return [];
    
    const ultimaFila = hojaDatos.getLastRow();
    if (ultimaFila <= 1) return [];
    
    // Obtenemos los valores y los valores formateados visibles en pantalla
    const rangoCompleto = hojaDatos.getRange(1, 1, ultimaFila, 16);
    const matrizValores = rangoCompleto.getValues();
    const matrizVisibles = rangoCompleto.getDisplayValues(); // Trae las fechas y números exactos como se ven
    
    const cabecera = matrizValores[0].map(h => h.toString().trim().toUpperCase());
    
    const idxFecha = cabecera.indexOf("FECHA");
    const idxProveedor = cabecera.indexOf("PROVEEDOR");
    const idxTipo = cabecera.indexOf("TIPO");
    const idxDocu = cabecera.indexOf("DOCU");
    const idxCompra = cabecera.indexOf("COMPRA");
    const idxPago = cabecera.indexOf("PAGO");
    const idxDias = cabecera.indexOf("DIAS");
    const idxVence = cabecera.indexOf("VENCE");
    const idxTipoPago = cabecera.indexOf("TIPO PAGO");
    const idxTasa = cabecera.indexOf("TASA");
    const idxBcv = cabecera.indexOf("BCV");
    const idxBolivares = cabecera.indexOf("BOLIVARES");
    const idxPagoBcv = cabecera.indexOf("PAGO A BCV");
    const idxPagoTasa = cabecera.indexOf("PAGO TASA");
    const idxDif = cabecera.indexOf("DIFERENCIA");
    const idxObs = cabecera.indexOf("OBSERVACIONES");
    
    const registros = [];
    const tz = Session.getScriptTimeZone();

    for (let i = 1; i < matrizValores.length; i++) {
      let filaValores = matrizValores[i];
      let filaVisibles = matrizVisibles[i];
      
      // Si la fila completa está totalmente en blanco, la omitimos
      if (!filaValores.join("").trim()) continue;
      
      // Asegurar que el nombre del proveedor exista
      let provNombre = (filaValores[idxProveedor] || "").toString().trim();
      if (!provNombre) continue; 
      
      // Formatear Fecha de Emisión de forma segura basada en el valor de pantalla
      let fechaEstandar = "";
      let visibleFecha = filaVisibles[idxFecha].trim();
      if (visibleFecha) {
        if (visibleFecha.includes('/')) {
          let partes = visibleFecha.split('/');
          if (partes.length === 3) {
            // Manejar formatos d/m/yyyy y dd/mm/yyyy
            let dia = partes[0].padStart(2, '0');
            let mes = partes[1].padStart(2, '0');
            let anio = partes[2].length === 2 ? "20" + partes[2] : partes[2];
            fechaEstandar = `${anio}-${mes}-${dia}`;
          }
        } else if (visibleFecha.includes('-')) {
          let partes = visibleFecha.split('-');
          if (partes[0].length === 4) {
            fechaEstandar = visibleFecha.substring(0, 10);
          } else if (partes.length === 3) {
            let dia = partes[0].padStart(2, '0');
            let mes = partes[1].padStart(2, '0');
            let anio = partes[2].length === 2 ? "20" + partes[2] : partes[2];
            fechaEstandar = `${anio}-${mes}-${dia}`;
          }
        }
      }
      
      // Si falla la conversión del valor visible, usamos la fecha cruda del objeto Date
      if (!fechaEstandar && filaValores[idxFecha] instanceof Date) {
        fechaEstandar = Utilities.formatDate(filaValores[idxFecha], tz, "yyyy-MM-dd");
      }
      
      // Formatear Fecha de Vencimiento de forma segura
      let venceEstandar = "";
      let visibleVence = filaVisibles[idxVence].trim();
      if (visibleVence) {
        if (visibleVence.includes('/')) {
          let p = visibleVence.split('/');
          if (p.length === 3) {
            let dia = p[0].padStart(2, '0');
            let mes = p[1].padStart(2, '0');
            let anio = p[2].length === 2 ? "20" + p[2] : p[2];
            venceEstandar = `${anio}-${mes}-${dia}`;
          }
        } else if (visibleVence.includes('-')) {
          let p = visibleVence.split('-');
          if (p[0].length === 4) {
            venceEstandar = visibleVence.substring(0, 10);
          } else if (p.length === 3) {
            let dia = p[0].padStart(2, '0');
            let mes = p[1].padStart(2, '0');
            let anio = p[2].length === 2 ? "20" + p[2] : p[2];
            venceEstandar = `${anio}-${mes}-${dia}`;
          }
        }
      }
      
      if (!venceEstandar && filaValores[idxVence] instanceof Date) {
        venceEstandar = Utilities.formatDate(filaValores[idxVence], tz, "yyyy-MM-dd");
      }
      
      registros.push({
        rowNum: i + 1,
        fecha: fechaEstandar,
        proveedor: provNombre,
        tipo: (filaValores[idxTipo] || "").toString().trim().toUpperCase(),
        docu: (filaValores[idxDocu] || "").toString().trim(),
        compra: parseFloat(filaValores[idxCompra]) || 0,
        pago: parseFloat(filaValores[idxPago]) || 0,
        dias: parseInt(filaValores[idxDias]) || 0,
        vence: venceEstandar,
        tipoPago: (filaValores[idxTipoPago] || "").toString().trim(),
        tasa: parseFloat(filaValores[idxTasa]) || 0,
        bcv: parseFloat(filaValores[idxBcv]) || 0,
        bolivares: parseFloat(filaValores[idxBolivares]) || 0,
        pagoBcv: parseFloat(filaValores[idxPagoBcv]) || 0,
        pagoTasa: parseFloat(filaValores[idxPagoTasa]) || 0,
        diferencia: parseFloat(filaValores[idxDif]) || 0,
        observacion: (filaValores[idxObs] || "").toString().trim()
      });
    }
    
    return registros.reverse();
  } catch (e) {
    Logger.log("Error en obtenerDatosProveedoresBackend: " + e.toString());
    return [];
  }
}

/**
 * CARGA DE CATÁLOGOS DESDE LAS HOJAS LOCALES del SPREADSHEET
 */
function obtenerListasFormularioBackend() {
  try {
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    let proveedores = [];
    let hProv = libro.getSheetByName("PROVEEDORES");
    if (hProv && hProv.getLastRow() >= 2) {
      proveedores = hProv.getRange(2, 1, hProv.getLastRow() - 1, 1).getValues().map(r => r[0].toString().trim()).filter(Boolean);
    }
    
    let tasaRef = 0;
    let bcvRef = 0;
    let hTasa = libro.getSheetByName("TASA");
    if (hTasa && hTasa.getLastRow() >= 2) {
      tasaRef = parseFloat(hTasa.getRange(2, 1).getValue()) || 0;
      bcvRef = parseFloat(hTasa.getRange(2, 2).getValue()) || 0;
    }
    
    let tiposPago = [];
    let hPago = libro.getSheetByName("T-PAGO");
    if (hPago && hPago.getLastRow() >= 2) {
      tiposPago = hPago.getRange(2, 1, hPago.getLastRow() - 1, 1).getValues().map(r => r[0].toString().trim()).filter(Boolean);
    }
    
    return {
      proveedores: proveedores,
      tasa: tasaRef,
      bcv: bcvRef,
      tiposPago: tiposPago
    };
  } catch (e) {
    return { proveedores: [], tasa: 0, bcv: 0, tiposPago: [] };
  }
}

function registrarOperacionProveedorBackend(objetoDatos) {
  try {
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hojaDatos = libro.getSheetByName("DATOS");
    if (!hojaDatos) return { success: false, message: "No se encontró la pestaña 'DATOS'." };
    
    let fechaFormateada = new Date(objetoDatos.fecha);
    fechaFormateada.setMinutes(fechaFormateada.getMinutes() + fechaFormateada.getTimezoneOffset());
    
    let venceFormateada = "";
    if (objetoDatos.vence) {
      venceFormateada = new Date(objetoDatos.vence);
      venceFormateada.setMinutes(venceFormateada.getMinutes() + venceFormateada.getTimezoneOffset());
    }
    
    const nuevaFila = [
      fechaFormateada,
      objetoDatos.proveedor,
      objetoDatos.tipo,
      objetoDatos.docu,
      objetoDatos.compra,
      objetoDatos.pago,
      objetoDatos.dias,
      venceFormateada,
      objetoDatos.tipoPago,
      objetoDatos.tasa,
      objetoDatos.bcv,
      objetoDatos.bolivares,
      objetoDatos.pagoBcv,
      objetoDatos.pagoTasa,
      objetoDatos.diferencia,
      objetoDatos.observacion
    ];
    
    hojaDatos.appendRow(nuevaFila);
    return { success: true, message: "Operación registrada exitosamente." };
  } catch (error) {
    return { success: false, message: "Error al guardar: " + error.toString() };
  }
}

function editarOperacionProveedorBackend(rowNum, objetoDatos) {
  try {
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hojaDatos = libro.getSheetByName("DATOS");
    if (!hojaDatos) return { success: false, message: "No se encontró la pestaña 'DATOS'." };
    
    let fechaFormateada = new Date(objetoDatos.fecha);
    fechaFormateada.setMinutes(fechaFormateada.getMinutes() + fechaFormateada.getTimezoneOffset());
    
    let venceFormateada = "";
    if (objetoDatos.vence) {
      venceFormateada = new Date(objetoDatos.vence);
      venceFormateada.setMinutes(venceFormateada.getMinutes() + venceFormateada.getTimezoneOffset());
    }
    
    const rangoFila = hojaDatos.getRange(rowNum, 1, 1, 16);
    rangoFila.setValues([[
      fechaFormateada,
      objetoDatos.proveedor,
      objetoDatos.tipo,
      objetoDatos.docu,
      objetoDatos.compra,
      objetoDatos.pago,
      objetoDatos.dias,
      venceFormateada,
      objetoDatos.tipoPago,
      objetoDatos.tasa,
      objetoDatos.bcv,
      objetoDatos.bolivares,
      objetoDatos.pagoBcv,
      objetoDatos.pagoTasa,
      objetoDatos.diferencia,
      objetoDatos.observacion
    ]]);
    
    return { success: true, message: "Registro actualizado correctamente." };
  } catch (error) {
    return { success: false, message: "Error al editar: " + error.toString() };
  }
}

function eliminarOperacionProveedorBackend(rowNum) {
  try {
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hojaDatos = libro.getSheetByName("DATOS");
    if (!hojaDatos) return { success: false, message: "No se encontró la pestaña 'DATOS'." };
    
    hojaDatos.deleteRow(rowNum);
    return { success: true, message: "Registro eliminado correctamente." };
  } catch (error) {
    return { success: false, message: "Error al eliminar: " + error.toString() };
  }
}
