/**
 * Servidor: Código.gs (Módulo de Proveedores V2)
 * Sistema de Control de Proveedores - Repuestos VZLA
 */

const CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQMd7oOMFdMuEWG5nvflNSptgLu1paKr5znNdaqCqoM8svH2weVz4-Zgyo8twnIUTrnjF7lvuap8x0t/pub?output=csv";

function doGet() {
  return HtmlService.createTemplateFromFile('Login')
      .evaluate()
      .setTitle('Repuestos VZLA - Control de Proveedores')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * VALIDACIÓN DE USUARIOS CONTRA LA PESTAÑA LOCAL 'USUARIOS'
 */
function iniciarSesion(usuario, contrasena) {
  try {
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hojaUsuarios = libro.getSheetByName("USUARIOS");
    
    if (!hojaUsuarios) {
      if (usuario.toLowerCase() === "admin" && contrasena === "1234") {
        return { success: true, usuario: "ADMIN", permiso: "TOTAL" };
      }
      return { success: false, message: "Error: No se encontró la pestaña 'USUARIOS'." };
    }
    
    const ultimaFila = hojaUsuarios.getLastRow();
    const datos = hojaUsuarios.getRange(2, 1, ultimaFila - 1, 3).getValues();
    
    for (let i = 0; i < datos.length; i++) {
      let usuarioHoja = datos[i][0].toString().trim();
      let contrasenaHoja = datos[i][1].toString().trim();
      let permisoHoja = datos[i][2].toString().trim().toUpperCase(); 
      
      if (usuario.toLowerCase().trim() === usuarioHoja.toLowerCase() && contrasena.trim() === contrasenaHoja) {
        return { success: true, usuario: usuarioHoja.toUpperCase(), permiso: permisoHoja };
      }
    }
    return { success: false, message: "Usuario o contraseña incorrectos." };
  } catch (error) {
    return { success: false, message: "Error en login: " + error.toString() };
  }
}

/**
 * OBTIENE LOS DATOS DEL PROVEEDOR DESDE LA URL PUBLICADA EN CSV
 */
function obtenerDatosProveedoresBackend(usuarioActual, permisoActual) {
  try {
    const respuesta = UrlFetchApp.fetch(CSV_URL);
    const contenidoCsv = respuesta.getContentText();
    const matriz = Utilities.parseCsv(contenidoCsv);
    
    if (matriz.length <= 1) return [];
    
    const cabecera = matriz[0].map(h => h.toString().trim().toUpperCase());
    
    const idxFecha     = cabecera.indexOf("FECHA");
    const idxProveedor = cabecera.indexOf("PROVEEDOR");
    const idxTipo      = cabecera.indexOf("TIPO");
    const idxDocu      = cabecera.indexOf("DOCU");
    const idxCompra    = cabecera.indexOf("COMPRA");
    const idxPago      = cabecera.indexOf("PAGO");
    const idxDias      = cabecera.indexOf("DIAS");
    const idxVence     = cabecera.indexOf("VENCE");
    const idxTipoPago  = cabecera.indexOf("TIPO PAGO");
    const idxTasa      = cabecera.indexOf("TASA");
    const idxBcv       = cabecera.indexOf("BCV");
    const idxBolivares = cabecera.indexOf("BOLIVARES");
    const idxPagoBcv   = cabecera.indexOf("PAGO A BCV");
    const idxPagoTasa  = cabecera.indexOf("PAGO TASA");
    const idxDif       = cabecera.indexOf("DIFERENCIA");
    const idxObs       = cabecera.indexOf("OBSERVACIONES");

    const registros = [];

    for (let i = 1; i < matriz.length; i++) {
      let fila = matriz[i];
      if (!fila[idxProveedor]) continue; 
      
      registros.push({
        rowNum: i + 1, // Guardamos el número de fila real en Sheets para poder editar/eliminar
        fecha: fila[idxFecha] || "",
        proveedor: fila[idxProveedor].toString().trim(),
        tipo: fila[idxTipo] || "",
        docu: fila[idxDocu] || "",
        compra: parseFloat(fila[idxCompra]) || 0,
        pago: parseFloat(fila[idxPago]) || 0,
        dias: parseInt(fila[idxDias]) || 0,
        vence: fila[idxVence] || "",
        tipoPago: fila[idxTipoPago] || "",
        tasa: parseFloat(fila[idxTasa]) || 0,
        bcv: parseFloat(fila[idxBcv]) || 0,
        bolivares: parseFloat(fila[idxBolivares]) || 0,
        pagoBcv: parseFloat(fila[idxPagoBcv]) || 0,
        pagoTasa: parseFloat(fila[idxPagoTasa]) || 0,
        diferencia: parseFloat(fila[idxDif]) || 0,
        observacion: fila[idxObs] || ""
      });
    }
    
    return registros.reverse(); 
  } catch (e) {
    return [];
  }
}

/**
 * CARGA DE CATÁLOGOS DESDE LAS HOJAS LOCALES DEL SPREADSHEET ACTIVO
 */
function obtenerListasFormularioBackend() {
  try {
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    
    let proveedores = [];
    let hProv = libro.getSheetByName("PROVEEDORES");
    if (hProv) {
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
    if (hPago) {
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

/**
 * REGISTRA UNA NUEVA FILA EN LA PESTAÑA 'DATOS'
 */
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

/**
 * ACTUALIZA UN REGISTRO EXISTENTE POR SU NÚMERO DE FILA
 */
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

    // Actualizamos el rango correspondiente de la fila de forma ordenada de la A a la P
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

/**
 * ELIMINA UN REGISTRO POR SU NÚMERO DE FILA
 */
function eliminarOperacionProveedorBackend(rowNum) {
  try {
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hojaDatos = libro.getSheetByName("DATOS");
    if (!hojaDatos) return { success: false, message: "No se encontró la pestaña 'DATOS'." };
    
    hojaDatos.deleteRow(rowNum);
    return { success: true, message: "Registro eliminado correctamente de la base de datos." };
  } catch (error) {
    return { success: false, message: "Error al eliminar: " + error.toString() };
  }
}

