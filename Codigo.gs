/**
 * ----------------------------------------------------------------------
 * HOLTMONT WORKSPACE V20 - BACKEND ROBUSTO (ScriptMaster Edition)
 * ----------------------------------------------------------------------
 */

// Referencia global al libro activo
const SS = SpreadsheetApp.getActiveSpreadsheet();

// --- CONFIGURACIÓN DE SEGURIDAD ---
const APP_CONFIG = {
  adminPass: "admin2025",  // <-- CAMBIA ESTO INMEDIATAMENTE
  userPass: "equipo2025",   // <-- CAMBIA ESTO INMEDIATAMENTE
  folderIdUploads: "",     // ID de carpeta Drive (Opcional, dejar vacío si no se usa)
  ppcSheetName: "PLANEACION SEMANAL",
  salesSheetName: "Datos" 
};

/* =========================================
   SERVIR HTML
   ========================================= */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Holtmont Workspace V20')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* =========================================
   HELPER: BUSCADOR INTELIGENTE DE HOJAS
   (Resuelve el problema de "No detecta la hoja")
   ========================================= */
function findSheetSmart(name) {
  if (!name) return null;
  
  // 1. Intento directo (Más rápido)
  let sheet = SS.getSheetByName(name);
  if (sheet) return sheet;

  // 2. Intento "Fuzzy" (Ignora mayúsculas y espacios extra al inicio/final)
  const cleanName = String(name).trim().toUpperCase();
  const allSheets = SS.getSheets();
  
  for (let s of allSheets) {
    const sName = s.getName().trim().toUpperCase();
    if (sName === cleanName) {
      return s;
    }
  }
  
  return null;
}

/* =========================================
   API: AUTENTICACIÓN Y CONFIGURACIÓN
   ========================================= */

function apiLogin(password) {
  if (password === APP_CONFIG.adminPass) {
    return { success: true, role: 'ADMIN', name: 'Administrador' };
  } else if (password === APP_CONFIG.userPass) {
    return { success: true, role: 'USER', name: 'Colaborador' };
  }
  return { success: false, message: 'Contraseña incorrecta' };
}

function getSystemConfig(role) {
  // Lista Maestra de Personal (Debe coincidir con nombres de pestañas)
  const staffList = [
    { name: "LUIS CARLOS", dept: "ADMINISTRACION" },
    { name: "ANTONIO SALAZAR", dept: "ADMINISTRACION" },
    { name: "ANGEL SALINAS", dept: "PROYECTOS" },
    { name: "RICARDO MENDO", dept: "COMPRAS" },
    { name: "CARLOS MENDEZ", dept: "PROYECTOS" },
    { name: "REYNALDO GARCIA", dept: "CONSTRUCCION" },
    { name: "INGE OLIVO", dept: "CONSTRUCCION" },
    { name: "EDUARDO TERAN", dept: "HVAC" },
    { name: "ALEXIS TORRES", dept: "DISEÑO" },
    { name: "TERESA GARZA", dept: "ADMINISTRACION" },
    { name: "RAMIRO RODRIGUEZ", dept: "ELECTROMECANICA" },
    { name: "GUILLERMO DAMICO", dept: "DISEÑO" },
    { name: "RUBEN PESQUEDA", dept: "CALIDAD" },
    { name: "JUDITH ECHAVARRIA", dept: "DISEÑO" },
    { name: "GISELA DOMINGUEZ", dept: "ADMINISTRACION" },
    { name: "VANESSA DE LARA", dept: "ADMINISTRACION" },
    { name: "NELSON MALDONADO", dept: "PROYECTOS" },
    { name: "VICTOR ALMACEN", dept: "ALMACEN" },
    { name: "DIMAS RAMOS", dept: "EHS" },
    { name: "CITLALI GOMEZ", dept: "ADMINISTRACION" },
    { name: "AIMEE RAMIREZ", dept: "ADMINISTRACION" },
    { name: "EDGAR LOPEZ", dept: "PROYECTOS" },
    { name: "INGE GALLARDO", dept: "CONSTRUCCION" },
    { name: "JEHU MARTINEZ", dept: "ELECTROMECANICA" },
    { name: "MIGUEL GONZALEZ", dept: "HVAC" },
    { name: "ALICIA RIVERA", dept: "LIMPIEZA" },
    { name: "JUAN JOSE SANCHEZ", dept: "ELECTROMECANICA" },
    { name: "EDUARDO MANZANARES", dept: "HVAC" },
    { name: "SELENE BALDONADO", dept: "ADMINISTRACION" },
    { name: "ROLANDO MORENO", dept: "PROYECTOS" },
    { name: "ROCIO CASTRO", dept: "DISEÑO" },
    { name: "DANIA GONZALEZ", dept: "DISEÑO" },
    { name: "JUANY RODRIGUEZ", dept: "ADMINISTRACION" },
    { name: "ANTONIA PINEDA", dept: "ADMINISTRACION" },
    { name: "LAURA HUERTA", dept: "ADMINISTRACION" },
    { name: "LILIANA MARTINEZ", dept: "ADMINISTRACION" },
    { name: "DANIELA CASTRO", dept: "ADMINISTRACION" },
    { name: "EDUARDO BENITEZ", dept: "PROYECTOS" },
    { name: "ANTONIO CABRERA", dept: "VENTAS" }
  ];

  return {
    departments: {
      "CONSTRUCCION": { label: "Construcción", icon: "fa-hard-hat", color: "#ffc107" },
      "HVAC": { label: "HVAC", icon: "fa-fan", color: "#0dcaf0" },
      "ELECTROMECANICA": { label: "Electromecánica", icon: "fa-bolt", color: "#dc3545" },
      "DISEÑO": { label: "Diseño & Ing.", icon: "fa-drafting-compass", color: "#6610f2" },
      "PROYECTOS": { label: "Proyectos", icon: "fa-project-diagram", color: "#0d6efd" },
      "ADMINISTRACION": { label: "Administración", icon: "fa-briefcase", color: "#6c757d" },
      "COMPRAS": { label: "Compras", icon: "fa-shopping-cart", color: "#198754" },
      "EHS": { label: "Seguridad (EHS)", icon: "fa-shield-alt", color: "#20c997" },
      "CALIDAD": { label: "Calidad", icon: "fa-flask", color: "#9c36b5" },
      "ALMACEN": { label: "Almacén", icon: "fa-warehouse", color: "#f7941d" },
      "LIMPIEZA": { label: "Limpieza", icon: "fa-broom", color: "#54b4d3" },
      "VENTAS": { label: "Ventas", icon: "fa-handshake", color: "#e83e8c" }
    },
    staff: staffList,
    specialModules: [
      { id: "PPC_MASTER", label: "PPC Maestro", icon: "fa-tasks", color: "#fd7e14", type: "ppc_native" },
      { id: "SALES_MASTER", label: "Control Ventas", icon: "fa-chart-line", color: "#d63384", type: "sales_native" }
    ]
  };
}

/* =========================================
   API: STAFF TRACKER (LECTURA Y EDICIÓN)
   ========================================= */

// Función auxiliar para buscar encabezados en las primeras 50 filas
function findHeaderRow(values) {
  // Buscamos palabras clave únicas de tu formato: "FOLIO", "ALTA", "CONCEPTO"
  for (let i = 0; i < Math.min(50, values.length); i++) {
    const rowStr = values[i].map(c => String(c).toUpperCase().trim()).join("|");
    // La fila correcta debe tener FOLIO, CONCEPTO y ALTA (según tus CSV)
    if (rowStr.includes("FOLIO") && rowStr.includes("CONCEPTO") && rowStr.includes("ALTA")) {
      return i;
    }
  }
  // Fallback: Si no encuentra ALTA, busca al menos FOLIO y CONCEPTO
  for (let i = 0; i < Math.min(50, values.length); i++) {
    const rowStr = values[i].map(c => String(c).toUpperCase().trim()).join("|");
    if (rowStr.includes("FOLIO") && rowStr.includes("CONCEPTO")) {
      return i;
    }
  }
  return -1;
}

function apiFetchStaffTrackerData(personName) {
  try {
    // Usamos el buscador inteligente
    const sheet = findSheetSmart(personName);
    
    if (!sheet) {
      return { success: true, data: [], message: `No se encontró la hoja: '${personName}'. Verifique el nombre en las pestañas.` };
    }

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { success: true, data: [], message: "Hoja vacía." };

    const headerRowIndex = findHeaderRow(values);

    if (headerRowIndex === -1) {
      return { success: true, data: [], message: "No se encontró la estructura de tabla (Faltan columnas FOLIO/CONCEPTO)." };
    }

    const headers = values[headerRowIndex].map(h => String(h).toUpperCase().trim().replace(/[\n\r]/g, ''));
    const dataRows = values.slice(headerRowIndex + 1);

    // Mapeo EXACTO a tus columnas basado en tus CSV
    // Folio,ALTA,FECHA,Hora,Clasificacion,Concepto,Involucrados,Avance %,Fecha estimada de fin,Hora estimada de fin,Reloj,Restricciones,Prioridades,Riesgos,Fecha respuesta,Correo,Carpeta,Status
    const colMap = {
      id: headers.indexOf("FOLIO"),
      alta: headers.indexOf("ALTA"),
      fecha: headers.indexOf("FECHA"),
      hora: headers.indexOf("HORA"),
      clasificacion: headers.indexOf("CLASIFICACION"),
      concepto: headers.indexOf("CONCEPTO"),
      involucrados: headers.indexOf("INVOLUCRADOS"),
      avance: headers.indexOf("AVANCE %"), // Ojo con el espacio y símbolo
      fechaEstFin: headers.indexOf("FECHA ESTIMADA DE FIN"),
      horaEstFin: headers.indexOf("HORA ESTIMADA DE FIN"),
      reloj: headers.indexOf("RELOJ"),
      restricciones: headers.indexOf("RESTRICCIONES"),
      prioridades: headers.indexOf("PRIORIDADES"),
      riesgos: headers.indexOf("RIESGOS"),
      status: headers.indexOf("STATUS")
    };

    const tasks = dataRows.map(row => {
      // Si no tiene concepto ni folio, saltar (fila vacía)
      if (!row[colMap.concepto] && !row[colMap.id]) return null;

      const formatVal = (idx) => (idx > -1 && row[idx] !== undefined) ? row[idx] : "";
      
      const formatTime = (val) => {
        if (val instanceof Date) return Utilities.formatDate(val, SS.getSpreadsheetTimeZone(), "dd/MM/yyyy");
        return val;
      };
      
      const formatHour = (val) => {
         if (val instanceof Date) return Utilities.formatDate(val, SS.getSpreadsheetTimeZone(), "HH:mm");
         return val;
      };

      return {
        id: formatVal(colMap.id),
        alta: formatVal(colMap.alta),
        fecha: formatTime(formatVal(colMap.fecha)),
        hora: formatHour(formatVal(colMap.hora)),
        clasificacion: formatVal(colMap.clasificacion),
        concepto: formatVal(colMap.concepto),
        involucrados: formatVal(colMap.involucrados),
        avance: formatVal(colMap.avance),
        fechaEstFin: formatTime(formatVal(colMap.fechaEstFin)),
        horaEstFin: formatHour(formatVal(colMap.horaEstFin)),
        reloj: formatVal(colMap.reloj),
        restricciones: formatVal(colMap.restricciones),
        prioridades: formatVal(colMap.prioridades),
        riesgos: formatVal(colMap.riesgos),
        status: formatVal(colMap.status)
      };
    }).filter(t => t !== null);

    return { success: true, data: tasks.reverse() }; // Mostrar lo más reciente arriba

  } catch (e) {
    Logger.log(e);
    return { success: false, message: "Error interno: " + e.toString() };
  }
}

/**
 * ACTUALIZA UNA TAREA EN LA HOJA DEL EMPLEADO (WRITE BACK)
 */
function apiUpdateTask(personName, taskData) {
  const lock = LockService.getScriptLock();
  // Bloqueo para evitar colisiones de escritura (5 segundos)
  if (lock.tryLock(5000)) {
    try {
      const sheet = findSheetSmart(personName);
      if (!sheet) return { success: false, message: "Hoja no encontrada para guardar." };

      const values = sheet.getDataRange().getValues();
      const headerRowIndex = findHeaderRow(values);
      if (headerRowIndex === -1) return { success: false, message: "Estructura de tabla inválida." };

      const headers = values[headerRowIndex].map(h => String(h).toUpperCase().trim());
      
      // Mapeo Inverso: Nombre de Campo -> Indice de Columna
      const fieldToColMap = {
        'id': headers.indexOf("FOLIO"),
        'alta': headers.indexOf("ALTA"),
        'fecha': headers.indexOf("FECHA"),
        'hora': headers.indexOf("HORA"),
        'clasificacion': headers.indexOf("CLASIFICACION"),
        'concepto': headers.indexOf("CONCEPTO"),
        'involucrados': headers.indexOf("INVOLUCRADOS"),
        'avance': headers.indexOf("AVANCE %"),
        'fechaEstFin': headers.indexOf("FECHA ESTIMADA DE FIN"),
        'horaEstFin': headers.indexOf("HORA ESTIMADA DE FIN"),
        'reloj': headers.indexOf("RELOJ"),
        'restricciones': headers.indexOf("RESTRICCIONES"),
        'prioridades': headers.indexOf("PRIORIDADES"),
        'riesgos': headers.indexOf("RIESGOS"),
        'status': headers.indexOf("STATUS")
      };

      // Identificar fila por ID (Folio) y Concepto para seguridad
      const idCol = fieldToColMap['id'];
      let targetRow = -1;

      // Iteramos buscando coincidencia
      for (let i = headerRowIndex + 1; i < values.length; i++) {
        const rowId = String(values[i][idCol] || "");
        const targetId = String(taskData.id || "");
        
        // Si hay ID, usalo. Si no, intenta coincidir por concepto (más arriesgado pero necesario si no hay folio)
        if (targetId && rowId === targetId) {
          targetRow = i + 1; // +1 Base 1
          break;
        } else if (!targetId && !rowId && String(values[i][fieldToColMap['concepto']]) === String(taskData.concepto)) {
           // Caso: Nueva tarea sin ID aún o fila sin ID
           targetRow = i + 1;
           break;
        }
      }

      if (targetRow === -1) {
        // Opción: Si no existe, ¿creamos nueva fila? 
        // Por ahora retornamos error para edición, pero podrías cambiar a appendRow
        return { success: false, message: "No se encontró el registro original para actualizar." };
      }

      // Escribir celda por celda solo los campos mapeados
      for (const [key, colIdx] of Object.entries(fieldToColMap)) {
        if (colIdx > -1 && taskData[key] !== undefined) {
          sheet.getRange(targetRow, colIdx + 1).setValue(taskData[key]);
        }
      }

      return { success: true, message: "Guardado exitoso." };

    } catch (e) {
      return { success: false, message: e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "El servidor está ocupado. Intenta de nuevo." };
  }
}

/* =========================================
   API: PPC (PLANIFICACIÓN)
   ========================================= */

function apiFetchPPCData() {
  try {
    const sheet = findSheetSmart(APP_CONFIG.ppcSheetName);
    if (!sheet) return { success: false, message: "Hoja PPC no encontrada" };
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: [] }; 

    // Leemos las ultimas 200 filas
    const startRow = Math.max(2, lastRow - 200); 
    // Ajustar columnas según tu hoja PPC real. Asumo primeras 10 col
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 10).getValues(); 

    const mappedData = data.map(row => ({
      id: row[0] || 'S/F',
      especialidad: row[1],
      concepto: row[2],
      responsable: row[3],
      fechaAlta: row[4] instanceof Date ? row[4].toLocaleDateString() : row[4],
      horas: row[5],
      cumplimiento: (row[6] && String(row[6]).toUpperCase() === 'SI') ? 'SI' : 'NO',
      archivoUrl: row[7],
      comentarios: row[8],
      clasificacion: "N/A", 
      prioridad: "Media"    
    })).filter(r => r.concepto); // Filtro simple

    return { success: true, data: mappedData };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function apiSavePPCData(payload) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      let sheet = findSheetSmart(APP_CONFIG.ppcSheetName);
      if (!sheet) sheet = SS.insertSheet(APP_CONFIG.ppcSheetName);

      const id = "ID-" + Math.floor(Math.random() * 100000);
      const fecha = new Date();
      
      const newRow = [
        id,
        payload.especialidad,
        payload.concepto,
        payload.responsable,
        fecha, // Fecha Alta
        payload.horas,
        payload.cumplimiento,
        payload.archivoUrl,
        payload.comentarios,
        "" // Comentarios Previos vacio
      ];

      sheet.appendRow(newRow);
      return { success: true, message: "Actividad guardada correctamente" };
    } catch (e) {
      return { success: false, message: e.toString() };
    } finally {
      lock.releaseLock();
    }
  }
  return { success: false, message: "Servidor ocupado" };
}

/* =========================================
   API: VENTAS
   ========================================= */
function apiFetchSalesData() {
  try {
    let sheet = findSheetSmart(APP_CONFIG.salesSheetName);
    if(!sheet) return { success: true, data: [], headers: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [], headers: [] };

    const headers = data[0].map(h => String(h).trim());
    const rows = data.slice(1);
    
    const formatted = rows.map(r => {
      let obj = {}; headers.forEach((h, i) => { obj[h] = r[i]; }); return obj;
    });

    return { success: true, data: formatted, headers: headers };

  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function apiSaveSaleData(jsonRow) {
  const lock = LockService.getScriptLock();
  if(lock.tryLock(5000)) {
    try {
      let sheet = findSheetSmart(APP_CONFIG.salesSheetName);
      if(!sheet) return { success: false, message: "Hoja ventas no existe" };
      
      const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      const newRow = headers.map(h => jsonRow[h] || "");
      
      sheet.appendRow(newRow);
      return { success: true };
    } catch(e) {
      return { success: false, message: e.toString() };
    } finally {
      lock.releaseLock();
    }
  }
}

/* =========================================
   UTILIDADES
   ========================================= */

function uploadFileToDrive(base64Data, type, fileName) {
  try {
    const folderId = APP_CONFIG.folderIdUploads; 
    let folder;
    if (folderId) {
      try { folder = DriveApp.getFolderById(folderId); } catch(e) { folder = DriveApp.getRootFolder(); }
    } else {
      folder = DriveApp.getRootFolder(); 
    }

    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf('base64,')+7));
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { success: true, fileUrl: file.getUrl() };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
