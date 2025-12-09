/**
 * ======================================================================
 * HOLTMONT WORKSPACE V126 - FIX VENTAS DISTRIBUTION
 * ======================================================================
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();

// --- CONFIGURACIÓN ---
const APP_CONFIG = {
  folderIdUploads: "", 
  ppcSheetName: "PPCV3",          // Historial
  draftSheetName: "PPC_BORRADOR", // Borrador persistente
  salesSheetName: "Datos",
  logSheetName: "LOG_SISTEMA"
};

// USUARIOS
const USER_DB = {
  "LUIS_CARLOS":    { pass: "admin2025", role: "ADMIN", label: "Administrador" },
  "JESUS_GARZA":    { pass: "ppc2025",   role: "PPC_ADMIN", label: "PPC Manager" },
  "ANTONIA_VENTAS": { pass: "tonita2025", role: "TONITA", label: "Ventas" },
  "JAIME_OLIVO":    { pass: "admin2025", role: "ADMIN_CONTROL", label: "Jaime Olivo" },
  "ANGEL_SALINAS":  { pass: "angel2025", role: "ANGEL_USER", label: "Angel Salinas" }
};

/* SERVICIO HTML */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Holtmont Workspace')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* HELPERS */
function findSheetSmart(name) {
  if (!name) return null;
  let sheet = SS.getSheetByName(name);
  if (sheet) return sheet;
  const clean = String(name).trim().toUpperCase();
  const all = SS.getSheets();
  for (let s of all) { if (s.getName().trim().toUpperCase() === clean) return s; }
  return null;
}

// DETECTOR ESTRICTO (NORMALIZADO)
function findHeaderRow(values) {
  for (let i = 0; i < Math.min(100, values.length); i++) {
    const rowStr = values[i].map(c => String(c).toUpperCase().replace(/\n/g, " ").replace(/\s+/g, " ").trim()).join("|");
    if (rowStr.includes("FOLIO") && rowStr.includes("CONCEPTO") && (rowStr.includes("ALTA") || rowStr.includes("AVANCE"))) return i;
    if (rowStr.includes("ID") && rowStr.includes("RESPONSABLE") && (rowStr.includes("FECHA") || rowStr.includes("ALTA"))) return i;
    if (rowStr.includes("CLIENTE") && (rowStr.includes("VENDEDOR") || rowStr.includes("AREA"))) return i;
  }
  return -1;
}

function logSystemEvent(user, action, details) {
  try {
    let sheet = SS.getSheetByName(APP_CONFIG.logSheetName);
    if (!sheet) {
      sheet = SS.insertSheet(APP_CONFIG.logSheetName);
      sheet.appendRow(["FECHA", "USUARIO", "ACCION", "DETALLES"]);
    }
    sheet.appendRow([new Date(), user, action, details]);
  } catch (e) { console.error(e); }
}

/* LOGIN */
function apiLogin(username, password) {
  const userKey = String(username).trim().toUpperCase();
  const user = USER_DB[userKey];
  if (user && user.pass === password) {
    logSystemEvent(userKey, "LOGIN", `Acceso exitoso (${user.role})`);
    return { success: true, role: user.role, name: user.label, username: userKey };
  }
  logSystemEvent(userKey || "ANONIMO", "LOGIN_FAIL", "Credenciales incorrectas");
  return { success: false, message: 'Usuario o contraseña incorrectos.' };
}

function getSystemConfig(role) {
  const fullDirectory = [
    { name: "ANTONIA_VENTAS", dept: "VENTAS" }, 
    { name: "JUDITH ECHAVARRIA", dept: "VENTAS" },
    { name: "EDUARDO MANZANARES", dept: "VENTAS" },
    { name: "RAMIRO RODRIGUEZ", dept: "VENTAS" },
    { name: "SEBASTIAN PADILLA", dept: "VENTAS" },
    { name: "CESAR GOMEZ", dept: "VENTAS" },
    { name: "ALFONSO CORREA", dept: "VENTAS" },
    { name: "TERESA GARZA", dept: "VENTAS" },
    { name: "GUILLERMO DAMICO", dept: "VENTAS" },
    { name: "ANGEL SALINAS", dept: "VENTAS" },
    { name: "JUAN JOSE SANCHEZ", dept: "VENTAS" },
    { name: "LUIS CARLOS", dept: "ADMINISTRACION" },
    { name: "ANTONIO SALAZAR", dept: "ADMINISTRACION" },
    { name: "ROCIO CASTRO", dept: "ADMINISTRACION" },
    { name: "DANIA GONZALEZ", dept: "ADMINISTRACION" },
    { name: "JUANY RODRIGUEZ", dept: "ADMINISTRACION" },
    { name: "LAURA HUERTA", dept: "ADMINISTRACION" },
    { name: "LILIANA MARTINEZ", dept: "ADMINISTRACION" },
    { name: "DANIELA CASTRO", dept: "ADMINISTRACION" },
    { name: "EDUARDO BENITEZ", dept: "ADMINISTRACION" },
    { name: "ANTONIO CABRERA", dept: "ADMINISTRACION" },
    { name: "ADMINISTRADOR", dept: "ADMINISTRACION" }, 
    { name: "EDUARDO MANZANARES", dept: "HVAC" },
    { name: "JUAN JOSE SANCHEZ", dept: "HVAC" },
    { name: "SELENE BALDONADO", dept: "HVAC" },
    { name: "ROLANDO MORENO", dept: "HVAC" },
    { name: "MIGUEL GALLARDO", dept: "ELECTROMECANICA" },
    { name: "SEBASTIAN PADILLA", dept: "ELECTROMECANICA" },
    { name: "JEHU MARTINEZ", dept: "ELECTROMECANICA" },
    { name: "MIGUEL GONZALEZ", dept: "ELECTROMECANICA" },
    { name: "ALICIA RIVERA", dept: "ELECTROMECANICA" },
    { name: "RICARDO MENDO", dept: "CONSTRUCCION" },
    { name: "CARLOS MENDEZ", dept: "CONSTRUCCION" },
    { name: "REYNALDO GARCIA", dept: "CONSTRUCCION" },
    { name: "INGE OLIVO", dept: "CONSTRUCCION" },
    { name: "EDUARDO TERAN", dept: "CONSTRUCCION" },
    { name: "EDGAR HOLT", dept: "CONSTRUCCION" },
    { name: "ALEXIS TORRES", dept: "CONSTRUCCION" },
    { name: "TERESA GARZA", dept: "CONSTRUCCION" },
    { name: "RAMIRO RODRIGUEZ", dept: "CONSTRUCCION" },
    { name: "GUILLERMO DAMICO", dept: "CONSTRUCCION" },
    { name: "RUBEN PESQUEDA", dept: "CONSTRUCCION" },
    { name: "JUDITH ECHAVARRIA", dept: "COMPRAS" },
    { name: "GISELA DOMINGUEZ", dept: "COMPRAS" },
    { name: "VANESSA DE LARA", dept: "COMPRAS" },
    { name: "NELSON MALDONADO", dept: "COMPRAS" },
    { name: "VICTOR ALMACEN", dept: "COMPRAS" }, 
    { name: "DIMAS RAMOS", dept: "EHS" },
    { name: "CITLALI GOMEZ", dept: "EHS" },
    { name: "AIMEE RAMIREZ", dept: "EHS" },
    { name: "EDGAR HOLT", dept: "MAQUINARIA" },
    { name: "ALEXIS TORRES", dept: "MAQUINARIA" },
    { name: "ANGEL SALINAS", dept: "DISEÑO" },
    { name: "EDGAR HOLT", dept: "DISEÑO" },
    { name: "EDGAR LOPEZ", dept: "DISEÑO" }
  ];
  const allDepts = {
      "CONSTRUCCION": { label: "Construcción", icon: "fa-hard-hat", color: "#e83e8c" },
      "COMPRAS": { label: "Compras/Almacén", icon: "fa-shopping-cart", color: "#198754" },
      "EHS": { label: "Seguridad (EHS)", icon: "fa-shield-alt", color: "#dc3545" },
      "DISEÑO": { label: "Diseño & Ing.", icon: "fa-drafting-compass", color: "#0d6efd" },
      "ELECTROMECANICA": { label: "Electromecánica", icon: "fa-bolt", color: "#ffc107" },
      "HVAC": { label: "HVAC", icon: "fa-fan", color: "#fd7e14" },
      "ADMINISTRACION": { label: "Administración", icon: "fa-briefcase", color: "#6f42c1" },
      "VENTAS": { label: "Ventas", icon: "fa-handshake", color: "#0dcaf0" },
      "MAQUINARIA": { label: "Maquinaria", icon: "fa-truck", color: "#20c997" }
  };
  
  if (role === 'TONITA') return { departments: { "VENTAS": allDepts["VENTAS"] }, allDepartments: allDepts, staff: [ { name: "ANTONIA_VENTAS", dept: "VENTAS" } ], directory: fullDirectory, specialModules: [] };
  
  const ppcModules = [
      { id: "PPC_MASTER", label: "PPC Maestro", icon: "fa-tasks", color: "#fd7e14", type: "ppc_native" },
      { id: "WEEKLY_PLAN", label: "Planeación Semanal", icon: "fa-calendar-alt", color: "#6f42c1", type: "weekly_plan_view" }
  ];

  if (role === 'PPC_ADMIN') return { departments: {}, allDepartments: allDepts, staff: [], directory: fullDirectory, specialModules: ppcModules };

  if (role === 'ADMIN_CONTROL') {
    return {
      departments: allDepts, allDepartments: allDepts, staff: fullDirectory, directory: fullDirectory,
      specialModules: [
        { id: "PPC_DINAMICO", label: "Tracker", icon: "fa-layer-group", color: "#e83e8c", type: "ppc_dynamic_view" },
        ...ppcModules,
        { id: "MIRROR_TONITA", label: "Monitor Toñita", icon: "fa-eye", color: "#0dcaf0", type: "mirror_staff", target: "ANTONIA_VENTAS" },
        { id: "ADMIN_TRACKER", label: "Control", icon: "fa-clipboard-list", color: "#6f42c1", type: "mirror_staff", target: "ADMINISTRADOR" }
      ]
    };
  }

  if (role === 'ANGEL_USER') {
    return {
      departments: { "DISEÑO": allDepts["DISEÑO"], "VENTAS": allDepts["VENTAS"] },
      allDepartments: allDepts, staff: [ { name: "ANGEL SALINAS", dept: "DISEÑO" } ], directory: fullDirectory,
      specialModules: [{ id: "MY_TRACKER", label: "Mi Tabla", icon: "fa-table", color: "#0d6efd", type: "mirror_staff", target: "ANGEL SALINAS" }]
    };
  }

  // ADMIN
  return {
    departments: allDepts, allDepartments: allDepts, staff: fullDirectory, directory: fullDirectory,
    specialModules: [ ...ppcModules, { id: "MIRROR_TONITA", label: "Monitor Toñita", icon: "fa-eye", color: "#0dcaf0", type: "mirror_staff", target: "ANTONIA_VENTAS" } ]
  };
}

/* 5. LECTURA DE DATOS STAFF */
function apiFetchStaffTrackerData(personName) {
  try {
    const sheet = findSheetSmart(personName);
    if (!sheet) return { success: true, data: [], history: [], headers: [], message: `Falta hoja: ${personName}` };
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { success: true, data: [], history: [], headers: [], message: "Vacía" };
    const headerRowIndex = findHeaderRow(values);
    if (headerRowIndex === -1) return { success: true, data: [], headers: [], message: "Sin formato válido" };
    const rawHeaders = values[headerRowIndex].map(h => String(h).trim());
    const validIndices = [];
    const cleanHeaders = [];
    rawHeaders.forEach((h, index) => {
      if(h !== "") { validIndices.push(index); cleanHeaders.push(h); }
    });
    const dataRows = values.slice(headerRowIndex + 1);
    const activeTasks = [];
    const historyTasks = [];
    let isReadingHistory = false;
    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      if (row.join("|").toUpperCase().includes("TAREAS REALIZADAS")) { isReadingHistory = true; continue; }
      if (row.every(c => c === "") || String(row[validIndices[0]]).toUpperCase() === String(cleanHeaders[0]).toUpperCase()) continue;
      let rowObj = {};
      let hasData = false;
      let sortDate = null;
      validIndices.forEach((colIndex, k) => {
        const headerName = cleanHeaders[k];
        let val = row[colIndex];
        if (val instanceof Date) {
           if (val.getFullYear() < 1900) val = Utilities.formatDate(val, SS.getSpreadsheetTimeZone(), "HH:mm");
           else {
              if (!sortDate) sortDate = val; 
              val = Utilities.formatDate(val, SS.getSpreadsheetTimeZone(), "dd/MM/yy");
           }
        } else if (typeof val === 'string') {
           if(val.match(/\d{1,2}\/\d{1,2}\/\d{4}/)) val = val.replace(/\/(\d{4})$/, (match, y) => "/" + y.slice(-2));
           else if (val.match(/\d{4}-\d{2}-\d{2}/)) { let d = new Date(val); val = Utilities.formatDate(d, SS.getSpreadsheetTimeZone(), "dd/MM/yy"); }
        }
        if (val !== "" && val !== undefined) hasData = true;
        rowObj[headerName] = val;
      });
      if (hasData) {
        rowObj['_sortDate'] = sortDate;
        rowObj['_rowIndex'] = headerRowIndex + i + 2;
        if (isReadingHistory) historyTasks.push(rowObj); else activeTasks.push(rowObj);
      }
    }
    const dateSorter = (a, b) => {
      const dA = a['_sortDate'] instanceof Date ? a['_sortDate'].getTime() : 0;
      const dB = b['_sortDate'] instanceof Date ? b['_sortDate'].getTime() : 0;
      return dB - dA;
    };
    return { success: true, data: activeTasks.sort(dateSorter).map(({_sortDate, ...rest}) => rest), history: historyTasks.sort(dateSorter).map(({_sortDate, ...rest}) => rest), headers: cleanHeaders };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function internalUpdateTask(personName, taskData) {
    try {
      const sheet = findSheetSmart(personName);
      if (!sheet) return { success: false, message: "Hoja no encontrada: " + personName };
      const values = sheet.getDataRange().getValues();
      const headerRowIndex = findHeaderRow(values);
      if (headerRowIndex === -1) return { success: false, message: "Sin tabla" };
      const sheetHeaders = values[headerRowIndex].map(h => String(h).toUpperCase().trim());
      const folioIndex = sheetHeaders.indexOf("FOLIO");
      const conceptoIndex = sheetHeaders.indexOf("CONCEPTO");
      
      const COL_MAP = {
        'ALTA': ['ALTA', 'CLIENTE', 'AREA', 'DEPARTAMENTO'],
        'INVOLUCRADOS': ['INVOLUCRADOS', 'VENDEDOR', 'RESPONSABLE', 'REQUISITOR'],
        'RELOJ': ['RELOJ', 'DIAS'],
        'FECHA': ['FECHA', 'FECHA INICIO', 'FECHA DE INICIO', 'FECHA ALTA'],
        'ESTATUS': ['ESTATUS', 'STATUS'],
        'CLASIFICACION': ['CLASIFICACION', 'CLASE', 'TIPO'], 
        'PRIORIDAD': ['PRIORIDAD', 'PRIORIDADES'],
        'AVANCE': ['AVANCE', 'AVANCE %'],
        'RESTRICCIONES': ['RESTRICCIONES', 'REST.'],
        'RIESGOS': ['RIESGOS'],
        'FECHA_RESPUESTA': ['FECHA RESPUESTA', 'FECHA DE RESPUESTA', 'HORA ESTIMADA DE FIN', 'FECHA ESTIMADA DE FIN', 'FEC. EST. FIN'],
        'HORA_ESTIMADA': ['HORA ESTIMADA DE FIN', 'HR. EST. FIN'],
        'F2': ['F2'], 'COTIZACION': ['COTIZACION', 'COT', 'COTIZACIÓN'], 'TIMEOUT': ['TIMEOUT', 'TIME OUT'], 'LAYOUT': ['LAYOUT'], 'TIMELINE': ['TIMELINE']
      };
      
      const getTargetColIdx = (key) => {
         const keyUpper = key.toUpperCase();
         let idx = sheetHeaders.indexOf(keyUpper);
         if (idx > -1) return idx;
         for (const std in COL_MAP) {
            if (keyUpper === std || COL_MAP[std].includes(keyUpper)) {
               for (const alias of COL_MAP[std]) {
                  const aliasIdx = sheetHeaders.indexOf(alias);
                  if (aliasIdx > -1) return aliasIdx;
               }
            }
         }
         return -1;
      };
      let targetRow = -1;
      if (taskData['_rowIndex']) { targetRow = parseInt(taskData['_rowIndex']); } 
      else {
          for (let i = headerRowIndex + 1; i < values.length; i++) {
            if (values[i].join("|").toUpperCase().includes("TAREAS REALIZADAS")) break;
            const rowValFolio = folioIndex > -1 ? String(values[i][folioIndex]).trim().toUpperCase() : "";
            const rowValConcepto = conceptoIndex > -1 ? String(values[i][conceptoIndex]).trim().toUpperCase() : "";
            const inFolio = String(taskData['FOLIO']||"").trim().toUpperCase();
            const inConcepto = String(taskData['CONCEPTO']||"").trim().toUpperCase();
            if (inFolio && inFolio === rowValFolio) { targetRow = i + 1; break; }
            else if (inConcepto && inConcepto === rowValConcepto) { targetRow = i + 1; break; }
          }
      }
      let isCompleted = false;
      const avanceKey = Object.keys(taskData).find(k => k.toUpperCase().includes("AVANCE"));
      if (avanceKey) { const val = String(taskData[avanceKey]).trim().replace('%', '');
        if (val === "100" || val === "1" || val === "1.0") isCompleted = true;
      }
      let result = { success: true, message: "Guardado" };
      if (isCompleted) {
        let rowValues;
        if (targetRow !== -1 && targetRow <= sheet.getLastRow()) { rowValues = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).getValues()[0]; } 
        else { rowValues = new Array(sheet.getLastColumn()).fill(""); }
        for (const key in taskData) {
           if (key.startsWith('_')) continue;
           const colIdx = getTargetColIdx(key);
           if (colIdx > -1) rowValues[colIdx] = taskData[key];
        }
        let historyIdx = -1;
        const allData = sheet.getDataRange().getValues();
        for(let r=0; r<allData.length; r++) { if(allData[r].join("|").toUpperCase().includes("TAREAS REALIZADAS")) { historyIdx = r+1; break; } }
        if (historyIdx === -1) { sheet.appendRow(["", "", "TAREAS REALIZADAS"]); sheet.appendRow(values[headerRowIndex]); }
        sheet.appendRow(rowValues);
        if (targetRow !== -1 && targetRow <= sheet.getLastRow()) { sheet.deleteRow(targetRow); }
        logSystemEvent(personName, "COMPLETE", `Archivado: ${taskData['CONCEPTO']}`);
        result = { success: true, message: "Tarea completada y movida.", moved: true };
      } 
      else if (targetRow !== -1) {
        for (const key in taskData) {
          if (key.startsWith('_')) continue;
          const colIdx = getTargetColIdx(key);
          if (colIdx > -1) sheet.getRange(targetRow, colIdx + 1).setValue(taskData[key]);
        }
        logSystemEvent(personName, "UPDATE", `Editado: ${taskData['CONCEPTO']}`);
        result = { success: true, message: "Guardado." };
      } else {
        const lastCol = sheet.getLastColumn();
        const headers = sheet.getRange(headerRowIndex+1, 1, 1, lastCol).getValues()[0].map(h => String(h).toUpperCase().trim());
        const newRow = headers.map(h => {
           const dataKey = Object.keys(taskData).find(k => {
              if (k.toUpperCase() === h) return true;
              for (const std in COL_MAP) { if (COL_MAP[std].includes(k.toUpperCase()) && COL_MAP[std].includes(h)) return true; }
              return false;
           });
           return dataKey ? taskData[dataKey] : "";
        });
        sheet.insertRowAfter(headerRowIndex + 1);
        sheet.getRange(headerRowIndex + 2, 1, 1, newRow.length).setValues([newRow]);
        logSystemEvent(personName, "CREATE", `Nuevo: ${taskData['CONCEPTO']}`);
        result = { success: true, message: "Creado." };
      }
      
      // LOGICA RESTAURADA: REPLICACION VENTAS -> VENDEDOR + ADMIN
      if (personName === "ANTONIA_VENTAS") {
          const distData = JSON.parse(JSON.stringify(taskData));
          delete distData._rowIndex; // Importante: borrar el indice para que se cree/busque de nuevo en destino
          
          // 1. ENVIAR AL VENDEDOR ASIGNADO
          const vendedorKey = Object.keys(taskData).find(k => k.toUpperCase() === "VENDEDOR");
          if (vendedorKey) {
              const vendedorName = taskData[vendedorKey];
              if (vendedorName && typeof vendedorName === 'string' && vendedorName.trim().length > 0) {
                  try { 
                      internalUpdateTask(vendedorName.trim(), distData); 
                  } catch(e) { 
                      console.warn("Error enviando a vendedor: " + e); 
                  }
              }
          }

          // 2. ENVIAR AL ADMINISTRADOR (CONTROL)
          try { 
             internalUpdateTask("ADMINISTRADOR", distData); 
          } catch(err) { 
             console.warn("Fallo replicación a ADMINISTRADOR: " + err); 
          }
      }
      return result;
    } catch(e) { return { success: false, message: e.toString() }; }
}

function apiUpdateTask(personName, taskData) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) { try { return internalUpdateTask(personName, taskData); } finally { lock.releaseLock(); } }
  return { success: false, message: "Ocupado." };
}

// --- GESTIÓN DE BORRADORES (PERSISTENCIA MEJORADA) ---
function apiFetchDrafts() {
  try {
    const sheet = findSheetSmart(APP_CONFIG.draftSheetName);
    if (!sheet) return { success: true, data: [] };
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 1) return { success: true, data: [] }; 
    const startRow = (rows[0][0] === "ESPECIALIDAD") ? 1 : 0;
    const drafts = rows.slice(startRow).map(r => ({
      especialidad: r[0], concepto: r[1], responsable: r[2], horas: r[3], cumplimiento: r[4],
      archivoUrl: r[5], comentarios: r[6], comentariosPrevios: r[7], 
      prioridades: r[8], riesgos: r[9], restricciones: r[10], fechaRespuesta: r[11], 
      clasificacion: r[12], fechaAlta: r[13] 
    })).filter(d => d.concepto);
    return { success: true, data: drafts };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function apiSyncDrafts(drafts) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      let sheet = findSheetSmart(APP_CONFIG.draftSheetName);
      if (!sheet) { sheet = SS.insertSheet(APP_CONFIG.draftSheetName); }
      sheet.clear();
      const headers = ["ESPECIALIDAD", "CONCEPTO", "RESPONSABLE", "HORAS", "CUMPLIMIENTO", "ARCHIVO", "COMENTARIOS", "PREVIOS", "PRIORIDAD", "RIESGOS", "RESTRICCIONES", "FECHA_RESP", "CLASIFICACION", "FECHA_ALTA"];
      if (drafts && drafts.length > 0) {
        const rows = drafts.map(d => [
          d.especialidad || "", d.concepto || "", d.responsable || "", d.horas || "", d.cumplimiento || "NO",
          d.archivoUrl || "", d.comentarios || "", d.comentariosPrevios || "",
          d.prioridades || "", d.riesgos || "", d.restricciones || "", d.fechaRespuesta || "", 
          d.clasificacion || "", d.fechaAlta || new Date() 
        ]);
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      } else {
        sheet.appendRow(headers);
      }
      return { success: true };
    } catch(e) { return { success: false, message: e.toString() }; } finally { lock.releaseLock(); }
  }
  return { success: false, message: "Ocupado syncing drafts" };
}

function apiClearDrafts() {
  try {
    const sheet = findSheetSmart(APP_CONFIG.draftSheetName);
    if(sheet) sheet.clear();
    return { success: true };
  } catch(e) { return { success: false }; }
}

function apiSavePPCData(payload) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const items = Array.isArray(payload) ? payload : [payload];
      let sheet = findSheetSmart(APP_CONFIG.ppcSheetName);
      
      if (!sheet) { 
        sheet = SS.insertSheet(APP_CONFIG.ppcSheetName); 
        sheet.appendRow(["ID", "Especialidad", "Descripción", "Responsable", "Fecha", "Reloj", "Cumplimiento", "Archivo", "Comentarios", "Comentarios Previos"]); 
      }
      
      const fechaHoy = new Date();
      const fechaStr = Utilities.formatDate(fechaHoy, SS.getSpreadsheetTimeZone(), "dd/MM/yy");
      
      items.forEach(item => {
          const id = "PPC-" + Math.floor(Math.random() * 100000);
          
          const lastRow = sheet.getLastRow();
          const nextRow = lastRow + 1;
          
          sheet.getRange(nextRow, 1, 1, 10).setValues([[
             id, item.especialidad, item.concepto, item.responsable, fechaHoy, 
             item.horas, item.cumplimiento, item.archivoUrl, item.comentarios, item.comentariosPrevios || ""
          ]]);

          const taskData = {
                 'FOLIO': id, 'CONCEPTO': item.concepto, 
                 'CLASIFICACION': item.clasificacion || "Media", 
                 'ALTA': item.especialidad, 'INVOLUCRADOS': item.responsable, 'FECHA': fechaStr,
                 'RELOJ': item.horas, 'ESTATUS': "ASIGNADO", 'PRIORIDAD': item.prioridad || item.prioridades, 
                 'RESTRICCIONES': item.restricciones, 'RIESGOS': item.riesgos, 
                 'FECHA_RESPUESTA': item.fechaRespuesta, 'AVANCE': "0%"
          };
          const responsables = String(item.responsable || "").split(",").map(s => s.trim()).filter(s => s);
          responsables.forEach(personName => {
            try { internalUpdateTask(personName, taskData); } catch(err) { console.warn(err); }
          });
          
          try { 
             internalUpdateTask("ADMINISTRADOR", taskData); 
          } catch(err) { 
             console.warn("Fallo replicación a ADMINISTRADOR: " + err); 
          }
      });
      return { success: true, message: "Procesado." };
    } catch (e) { return { success: false, message: e.toString() }; } finally { lock.releaseLock(); }
  }
  return { success: false, message: "Ocupado." };
}

function uploadFileToDrive(data, type, name) {
  try {
    const folderId = APP_CONFIG.folderIdUploads;
    let folder;
    if (folderId && folderId.trim() !== "") { try { folder = DriveApp.getFolderById(folderId); } catch(e) { folder = DriveApp.getRootFolder(); } } 
    else { folder = DriveApp.getRootFolder(); }
    const blob = Utilities.newBlob(Utilities.base64Decode(data.split(',')[1]), type, name);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, fileUrl: file.getUrl() };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// FUNCION CRITICA MEJORADA: LECTURA INTELIGENTE
function apiFetchPPCData() { 
  try { 
    const s = findSheetSmart(APP_CONFIG.ppcSheetName); 
    if(!s) return {success:true,data:[]}; 
    
    const range = s.getDataRange();
    const values = range.getValues();
    if (values.length < 2) return {success:true, data:[]};
    
    const headerIdx = findHeaderRow(values);
    if (headerIdx === -1) return {success:true, data:[]};

    const headers = values[headerIdx].map(h => String(h).toUpperCase().replace(/\n/g, " ").trim());
    const colMap = {
      id: headers.findIndex(h => h.includes("ID")),
      esp: headers.findIndex(h => h.includes("ESPECIALIDAD")),
      con: headers.findIndex(h => h.includes("DESCRIPCI")), 
      resp: headers.findIndex(h => h.includes("RESPONSABLE")),
      fecha: headers.findIndex(h => h.includes("FECHA") || h.includes("ALTA")),
      reloj: headers.findIndex(h => h.includes("RELOJ")),
      cump: headers.findIndex(h => h.includes("CUMPLIMIENTO")),
      arch: headers.findIndex(h => h.includes("ARCHIVO")),
      com: headers.findIndex(h => h.includes("COMENTARIOS") && h.includes("CURSO")),
      prev: headers.findIndex(h => h.includes("COMENTARIOS") && h.includes("PREVIA"))
    };

    let dataRows = values.slice(headerIdx + 1);
    if(dataRows.length > 300) dataRows = dataRows.slice(dataRows.length - 300);

    const resultData = dataRows.map(r => {
      const getVal = (idx) => (idx > -1 && r[idx] !== undefined) ? r[idx] : "";
      return {
        id: getVal(colMap.id),
        especialidad: getVal(colMap.esp),
        concepto: getVal(colMap.con),
        responsable: getVal(colMap.resp),
        fechaAlta: getVal(colMap.fecha),
        horas: getVal(colMap.reloj),
        cumplimiento: getVal(colMap.cump),
        archivoUrl: getVal(colMap.arch),
        comentarios: getVal(colMap.com),
        comentariosPrevios: getVal(colMap.prev)
      };
    }).filter(x => x.concepto).reverse();

    return { success: true, data: resultData }; 
  } catch(e){ return {success:false, message: e.toString()} } 
}

function apiFetchSalesData() { try { const s = findSheetSmart(APP_CONFIG.salesSheetName); if(!s) return {success:true,data:[],headers:[]};
const d = s.getDataRange().getValues(); return {success:true, headers:d[0], data:d.slice(1).map(r=>{let o={};d[0].forEach((h,i)=>o[h]=r[i]);return o;})}; } catch(e){return {success:false}} }
function apiSaveSaleData(j) { const s = findSheetSmart(APP_CONFIG.salesSheetName);
s.appendRow(s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(h=>j[h]||"")); return {success:true}; }
