/**
 * ======================================================================
 * HOLTMONT WORKSPACE V128 - SCRIPTMASTER EDITION
 * Optimizado: Batch Write, Smart Headers, KPI Dashboard Support
 * ======================================================================
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();

// --- CONFIGURACIÓN ---
const APP_CONFIG = {
  folderIdUploads: "", 
  ppcSheetName: "PPCV3",          // Historial Maestro / Planeación Semanal
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

// DETECTOR DE CABECERAS INTELIGENTE (ScriptMaster)
function findHeaderRow(values) {
  for (let i = 0; i < Math.min(100, values.length); i++) {
    const rowStr = values[i].map(c => String(c).toUpperCase().replace(/\n/g, " ").replace(/\s+/g, " ").trim()).join("|");
    
    // 1. Tracker Estandar
    if (rowStr.includes("FOLIO") && rowStr.includes("CONCEPTO")) return i;
    
    // 2. Staff Antiguo
    if (rowStr.includes("ID") && rowStr.includes("RESPONSABLE")) return i;
    
    // 3. CASO PPCV3 (Flexible): Acepta ID o FOLIO + DESCRIPCION o RESPONSABLE
    if ((rowStr.includes("FOLIO") || rowStr.includes("ID")) && 
        (rowStr.includes("DESCRIPCI") || rowStr.includes("RESPONSABLE") || rowStr.includes("CONCEPTO"))) {
      return i;
    }

    // 4. Ventas
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

/**
 * ======================================================================
 * OPTIMIZACIÓN SCRIPTMASTER: ESCRITURA QUIRÚRGICA E INTELIGENTE
 * Detecta si es un solo cambio para escribir solo esa fila.
 * ======================================================================
 */
function internalBatchUpdateTasks(sheetName, tasksArray) {
  if (!tasksArray || tasksArray.length === 0) return { success: true };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { success: false, message: "Hoja ocupada, intenta de nuevo." };

  try {
    const sheet = findSheetSmart(sheetName);
    if (!sheet) return { success: false, message: "Hoja no encontrada: " + sheetName };

    // Leemos datos para mapear
    const dataRange = sheet.getDataRange();
    let values = dataRange.getValues();
    if (values.length === 0) return { success: false, message: "Hoja vacía" };

    const headerRowIndex = findHeaderRow(values);
    if (headerRowIndex === -1) return { success: false, message: "Sin cabeceras válidas" };

    const headers = values[headerRowIndex].map(h => String(h).toUpperCase().trim());
    const colMap = {};
    headers.forEach((h, i) => colMap[h] = i);

    const getColIdx = (key) => {
      const k = key.toUpperCase();
      if (colMap[k] !== undefined) return colMap[k];
      const aliases = {
        'FECHA': ['FECHA', 'FECHA ALTA', 'FECHA INICIO'],
        'CONCEPTO': ['CONCEPTO', 'DESCRIPCION', 'DESCRIPCIÓN DE LA ACTIVIDAD'],
        'RESPONSABLE': ['RESPONSABLE', 'INVOLUCRADOS'],
        'RELOJ': ['RELOJ', 'HORAS', 'DIAS'],
        'ESTATUS': ['ESTATUS', 'STATUS'],
        'CUMPLIMIENTO': ['CUMPLIMIENTO', 'CUMPL.', 'CUMP'] // Alias crítico
      };
      for (let main in aliases) {
        if (aliases[main].includes(k)) {
             for(let alias of aliases[main]) if(colMap[alias] !== undefined) return colMap[alias];
        }
      }
      return -1;
    };

    const folioIdx = getColIdx('FOLIO') > -1 ? getColIdx('FOLIO') : getColIdx('ID');
    let singleRowIndex = -1; // Para optimización

    tasksArray.forEach(task => {
      let rowIndex = -1;

      // Usamos el índice de fila si viene del frontend (Más seguro y rápido)
      if (task._rowIndex) {
        rowIndex = parseInt(task._rowIndex) - 1; // Ajuste a base 0
      } else {
        // Búsqueda fallback por ID
        const tFolio = String(task['FOLIO'] || task['ID'] || "").toUpperCase();
        for (let i = headerRowIndex + 1; i < values.length; i++) {
           const row = values[i];
           if (folioIdx > -1 && String(row[folioIdx]).toUpperCase() === tFolio) { rowIndex = i; break; }
        }
      }

      if (rowIndex > -1 && rowIndex < values.length) {
        // UPDATE en memoria
        Object.keys(task).forEach(key => {
          if (key.startsWith('_')) return;
          const cIdx = getColIdx(key);
          if (cIdx > -1) values[rowIndex][cIdx] = task[key];
        });
        singleRowIndex = rowIndex;
      }
    });

    // --- ESCRITURA OPTIMIZADA (QUIRÚRGICA) ---
    // Si solo modificamos una fila, escribimos SOLO esa fila.
    if (tasksArray.length === 1 && singleRowIndex > -1) {
       // Escribimos en (FilaReal, Columna 1, 1 Fila, AnchoTotal)
       sheet.getRange(singleRowIndex + 1, 1, 1, values[0].length).setValues([values[singleRowIndex]]);
    } else {
       // Si son muchos cambios, escribimos todo el bloque (Batch tradicional)
       sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    }
    
    return { success: true };

  } catch (e) {
    console.error(e);
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// NUEVA FUNCIÓN PARA GUARDAR CAMBIOS DE PPCV3
function apiUpdatePPCV3(taskData) {
  return internalBatchUpdateTasks(APP_CONFIG.ppcSheetName, [taskData]);
}

// Wrapper para llamadas individuales
function internalUpdateTask(personName, taskData) {
    try {
        const res = internalBatchUpdateTasks(personName, [taskData]);
        if (personName === "ANTONIA_VENTAS") {
             const distData = JSON.parse(JSON.stringify(taskData));
             delete distData._rowIndex; 
             const vendedorKey = Object.keys(taskData).find(k => k.toUpperCase() === "VENDEDOR");
             if (vendedorKey && taskData[vendedorKey]) {
                 try { internalBatchUpdateTasks(taskData[vendedorKey].trim(), [distData]); } catch(e){}
             }
             try { internalBatchUpdateTasks("ADMINISTRADOR", [distData]); } catch(e){}
        }
        return res;
    } catch(e) { return {success:false, message:e.toString()}; }
}

function apiUpdateTask(personName, taskData) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) { 
      try { return internalUpdateTask(personName, taskData); } 
      finally { lock.releaseLock(); } 
  }
  return { success: false, message: "Ocupado." };
}

// --- GESTIÓN DE BORRADORES ---
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

// ======================================================================
// FUNCION CRITICA OPTIMIZADA
// REQUERIMIENTO: Enviar a Trabajador + Duplicar Admin + Guardar PPCV3
// ======================================================================
function apiSavePPCData(payload) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(20000)) { 
    try {
      const items = Array.isArray(payload) ? payload : [payload];
      
      let sheetPPC = findSheetSmart(APP_CONFIG.ppcSheetName);
      if (!sheetPPC) { 
        sheetPPC = SS.insertSheet(APP_CONFIG.ppcSheetName);
        sheetPPC.appendRow(["ID", "Especialidad", "Descripción", "Responsable", "Fecha", "Reloj", "Cumplimiento", "Archivo", "Comentarios", "Comentarios Previos"]);
      }
      
      const fechaHoy = new Date();
      const fechaStr = Utilities.formatDate(fechaHoy, SS.getSpreadsheetTimeZone(), "dd/MM/yy");
      
      const rowsForPPC = [];
      const tasksBySheet = {};
      const addTaskToSheet = (sheetName, task) => {
          if (!sheetName) return;
          const key = sheetName.trim();
          if (!tasksBySheet[key]) tasksBySheet[key] = [];
          tasksBySheet[key].push(task);
      };

      items.forEach(item => {
          const id = "PPC-" + Math.floor(Math.random() * 100000);
          
          rowsForPPC.push([
             id, 
             item.especialidad, item.concepto, item.responsable, fechaHoy, 
             item.horas, item.cumplimiento, item.archivoUrl, item.comentarios, 
             item.comentariosPrevios || ""
          ]);

          const taskData = {
                 'FOLIO': id, 
                 'CONCEPTO': item.concepto, 
                 'CLASIFICACION': item.clasificacion || "Media", 
                 'ALTA': item.especialidad, 
                 'INVOLUCRADOS': item.responsable, 
                 'FECHA': fechaStr,
                 'RELOJ': item.horas, 
                 'ESTATUS': "ASIGNADO", 
                 'PRIORIDAD': item.prioridad || item.prioridades, 
                 'RESTRICCIONES': item.restricciones, 
                 'RIESGOS': item.riesgos, 
                 'FECHA_RESPUESTA': item.fechaRespuesta, 
                 'AVANCE': "0%"
          };
          // DUPLICADO EN ADMIN
          addTaskToSheet("ADMINISTRADOR", taskData);
          // DISTRIBUCION
          const responsables = String(item.responsable || "").split(",").map(s => s.trim()).filter(s => s);
          responsables.forEach(personName => {
             addTaskToSheet(personName, taskData);
          });
      });

      // A. Guardar en PPCV3
      if (rowsForPPC.length > 0) {
          const lastRow = sheetPPC.getLastRow();
          sheetPPC.getRange(lastRow + 1, 1, rowsForPPC.length, rowsForPPC[0].length).setValues(rowsForPPC);
      }

      // B. Distribuir BATCH
      for (const [targetSheet, tasks] of Object.entries(tasksBySheet)) {
          internalBatchUpdateTasks(targetSheet, tasks);
      }

      return { success: true, message: "Procesado y Distribuido Correctamente." };
    } catch (e) { 
        console.error(e);
        return { success: false, message: e.toString() };
    } finally { 
        lock.releaseLock();
    }
  }
  return { success: false, message: "Sistema Ocupado, intenta de nuevo." };
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
      id: headers.findIndex(h => h.includes("ID") || h.includes("FOLIO")),
      esp: headers.findIndex(h => h.includes("ESPECIALIDAD")),
      con: headers.findIndex(h => h.includes("DESCRIPCI") || h.includes("CONCEPTO")), 
      resp: headers.findIndex(h => h.includes("RESPONSABLE") || h.includes("INVOLUCRADOS")),
      fecha: headers.findIndex(h => h.includes("FECHA") || h.includes("ALTA")),
      reloj: headers.findIndex(h => h.includes("RELOJ")),
      cump: headers.findIndex(h => h.includes("CUMPLIMIENTO")),
      arch: headers.findIndex(h => h.includes("ARCHIVO") || h.includes("CLIP")),
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

/**
 * LECTURA INTELIGENTE DE PPCV3 PARA PLANEACIÓN SEMANAL
 */
function apiFetchWeeklyPlanData() {
  try {
    const sheet = findSheetSmart(APP_CONFIG.ppcSheetName);
    if (!sheet) return { success: false, message: "No existe la hoja PPCV3" };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, headers: [], data: [] };

    const headerRowIdx = findHeaderRow(data);
    if (headerRowIdx === -1) return { success: false, message: "Cabeceras no encontradas en PPCV3." };

    const originalHeaders = data[headerRowIdx].map(h => String(h).trim());
    
    // Normalizar Cabeceras (Mapeo Inteligente)
    const mappedHeaders = originalHeaders.map(h => {
        const up = h.toUpperCase();
        if (up.includes("DESCRIPCI") || up.includes("CONCEPTO")) return "CONCEPTO"; 
        if (up.includes("INVOLUCRADOS") || up.includes("RESPONSABLE")) return "RESPONSABLE";
        if (up.includes("ALTA") || up.includes("FECHA")) return "FECHA";
        if (up.includes("RELOJ") || up.includes("HORAS")) return "RELOJ";
        if (up.includes("ARCHIV") || up.includes("CLIP")) return "ARCHIVO";
        if (up.includes("CUMPLIMIENTO")) return "CUMPLIMIENTO";
        return up; 
    });

    const displayHeaders = ["SEMANA", ...mappedHeaders];
    
    const rows = data.slice(headerRowIdx + 1);
    const result = rows.map((r, i) => {
      const rowObj = { _rowIndex: headerRowIdx + i + 2 };
      
      mappedHeaders.forEach((h, colIdx) => {
        let val = r[colIdx];
        if (val instanceof Date) {
           val = Utilities.formatDate(val, SS.getSpreadsheetTimeZone(), "dd/MM/yyyy");
        }
        rowObj[h] = val;
      });

      // Calculo de Semana
      const fechaVal = rowObj["FECHA"];
      let semanaNum = "-";
      if (fechaVal) {
        let dateObj = null;
        if (String(fechaVal).includes("/")) {
          const parts = String(fechaVal).split("/"); 
          if(parts.length === 3) dateObj = new Date(parts[2], parts[1]-1, parts[0]);
        } else if (fechaVal instanceof Date) {
          dateObj = fechaVal;
        } else {
          dateObj = new Date(fechaVal);
        }
        if (dateObj && !isNaN(dateObj.getTime())) semanaNum = getWeekNumber(dateObj); 
      }
      rowObj["SEMANA"] = semanaNum;
      
      return rowObj;
    }).filter(r => r["CONCEPTO"] || r["ID"] || r["FOLIO"]); 

    return { success: true, headers: displayHeaders, data: result.reverse() }; 
  } catch (e) {
    console.error(e);
    return { success: false, message: e.toString() };
  }
}

// Helper: Semana ISO
function getWeekNumber(d) {
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay()||7));
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
  var weekNo = Math.ceil(( ( (d - yearStart) / 86400000) + 1)/7);
  return weekNo;
}

function apiFetchSalesData() { try { const s = findSheetSmart(APP_CONFIG.salesSheetName);
if(!s) return {success:true,data:[],headers:[]};
const d = s.getDataRange().getValues(); return {success:true, headers:d[0], data:d.slice(1).map(r=>{let o={};d[0].forEach((h,i)=>o[h]=r[i]);return o;})};
} catch(e){return {success:false}} }
function apiSaveSaleData(j) { const s = findSheetSmart(APP_CONFIG.salesSheetName);
s.appendRow(s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(h=>j[h]||"")); return {success:true}; }
