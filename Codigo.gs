/**
 * ======================================================================
 * HOLTMONT WORKSPACE V86 - BACKEND FINAL (FULL)
 * ======================================================================
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();

// --- CONFIGURACIÓN ---
const APP_CONFIG = {
  folderIdUploads: "", 
  ppcSheetName: "PPCV3",
  salesSheetName: "Datos",
  logSheetName: "LOG_SISTEMA"
};

// USUARIOS
const USER_DB = {
  "LUIS_CARLOS": { pass: "admin2025", role: "ADMIN", label: "Administrador" },
  "JESUS_GARZA": { pass: "ppc2025",   role: "PPC_ADMIN", label: "PPC Manager" },
  "TONITA_MX":   { pass: "tonita2025", role: "TONITA", label: "Ventas" }
};

/* SERVICIO HTML */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Holtmont Workspace V86')
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

// DETECTOR ESTRICTO (Ignora tabla resumen)
function findHeaderRow(values) {
  for (let i = 0; i < Math.min(50, values.length); i++) {
    const rowStr = values[i].map(c => String(c).toUpperCase().trim()).join("|");
    // ADMIN (Busca ALTA o AVANCE)
    if (rowStr.includes("FOLIO") && rowStr.includes("CONCEPTO") && (rowStr.includes("ALTA") || rowStr.includes("AVANCE"))) return i;
    // VENTAS (Busca CLIENTE)
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
    { name: "TONITA_MX", dept: "VENTAS" },
    { name: "ANTONIO CABRERA", dept: "VENTAS" },
    { name: "LUIS CARLOS", dept: "ADMINISTRACION" },
    { name: "ANTONIO SALAZAR", dept: "ADMINISTRACION" },
    { name: "TERESA GARZA", dept: "ADMINISTRACION" },
    { name: "GISELA DOMINGUEZ", dept: "ADMINISTRACION" },
    { name: "VANESSA DE LARA", dept: "ADMINISTRACION" },
    { name: "CITLALI GOMEZ", dept: "ADMINISTRACION" },
    { name: "AIMEE RAMIREZ", dept: "ADMINISTRACION" },
    { name: "SELENE BALDONADO", dept: "ADMINISTRACION" },
    { name: "JUANY RODRIGUEZ", dept: "ADMINISTRACION" },
    { name: "ANTONIA PINEDA", dept: "ADMINISTRACION" },
    { name: "LAURA HUERTA", dept: "ADMINISTRACION" },
    { name: "LILIANA MARTINEZ", dept: "ADMINISTRACION" },
    { name: "DANIELA CASTRO", dept: "ADMINISTRACION" },
    { name: "ANGEL SALINAS", dept: "PROYECTOS" },
    { name: "CARLOS MENDEZ", dept: "PROYECTOS" },
    { name: "NELSON MALDONADO", dept: "PROYECTOS" },
    { name: "EDGAR LOPEZ", dept: "PROYECTOS" },
    { name: "ROLANDO MORENO", dept: "PROYECTOS" },
    { name: "EDUARDO BENITEZ", dept: "PROYECTOS" },
    { name: "RICARDO MENDO", dept: "COMPRAS" },
    { name: "REYNALDO GARCIA", dept: "CONSTRUCCION" },
    { name: "INGE OLIVO", dept: "CONSTRUCCION" },
    { name: "INGE GALLARDO", dept: "CONSTRUCCION" },
    { name: "EDUARDO TERAN", dept: "HVAC" },
    { name: "MIGUEL GONZALEZ", dept: "HVAC" },
    { name: "EDUARDO MANZANARES", dept: "HVAC" },
    { name: "ALEXIS TORRES", dept: "DISEÑO" },
    { name: "GUILLERMO DAMICO", dept: "DISEÑO" },
    { name: "JUDITH ECHAVARRIA", dept: "DISEÑO" },
    { name: "ROCIO CASTRO", dept: "DISEÑO" },
    { name: "DANIA GONZALEZ", dept: "DISEÑO" },
    { name: "RAMIRO RODRIGUEZ", dept: "ELECTROMECANICA" },
    { name: "JEHU MARTINEZ", dept: "ELECTROMECANICA" },
    { name: "JUAN JOSE SANCHEZ", dept: "ELECTROMECANICA" },
    { name: "RUBEN PESQUEDA", dept: "CALIDAD" },
    { name: "VICTOR ALMACEN", dept: "ALMACEN" },
    { name: "DIMAS RAMOS", dept: "EHS" },
    { name: "ALICIA RIVERA", dept: "LIMPIEZA" }
  ];

  const allDepts = {
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
  };

  if (role === 'TONITA') {
    return {
      departments: { "VENTAS": { label: "Mis Tablas", icon: "fa-chart-line", color: "#e83e8c" } },
      staff: [ { name: "TONITA_MX", dept: "VENTAS" } ],
      directory: fullDirectory,
      specialModules: [] 
    };
  }

  if (role === 'PPC_ADMIN') {
    return {
      departments: allDepts,
      staff: [],
      directory: fullDirectory,
      specialModules: [{ id: "PPC_MASTER", label: "PPC Maestro", icon: "fa-tasks", color: "#fd7e14", type: "ppc_native" }]
    };
  }

  return {
    departments: allDepts,
    staff: fullDirectory,
    directory: fullDirectory,
    specialModules: [
      { id: "PPC_MASTER", label: "PPC Maestro", icon: "fa-tasks", color: "#fd7e14", type: "ppc_native" },
      { id: "SALES_MASTER", label: "Control Ventas", icon: "fa-chart-line", color: "#d63384", type: "sales_native" },
      { id: "MIRROR_TONITA", label: "Monitor Toñita", icon: "fa-eye", color: "#6f42c1", type: "mirror_staff", target: "TONITA_MX" }
    ]
  };
}

/* 5. LECTURA DE DATOS */
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
      if (row.join("|").toUpperCase().includes("TAREAS REALIZADAS")) {
        isReadingHistory = true; continue;
      }
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
              val = Utilities.formatDate(val, SS.getSpreadsheetTimeZone(), "dd/MM/yyyy");
           }
        }
        if (val !== "" && val !== undefined) hasData = true;
        rowObj[headerName] = val;
      });
      
      if (hasData) {
        rowObj['_sortDate'] = sortDate;
        if (isReadingHistory) historyTasks.push(rowObj); else activeTasks.push(rowObj);
      }
    }

    const dateSorter = (a, b) => {
      const dA = a['_sortDate'] instanceof Date ? a['_sortDate'].getTime() : 0;
      const dB = b['_sortDate'] instanceof Date ? b['_sortDate'].getTime() : 0;
      return dB - dA;
    };

    return { 
      success: true, 
      data: activeTasks.sort(dateSorter).map(({_sortDate, ...rest}) => rest), 
      history: historyTasks.sort(dateSorter).map(({_sortDate, ...rest}) => rest),
      headers: cleanHeaders 
    };
  } catch (e) { return { success: false, message: e.toString() }; }
}

/* 6. ESCRITURA + MOVER 100% */
function apiUpdateTask(personName, taskData) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      const sheet = findSheetSmart(personName);
      if (!sheet) return { success: false, message: "Hoja no encontrada" };
      
      const values = sheet.getDataRange().getValues();
      const headerRowIndex = findHeaderRow(values);
      if (headerRowIndex === -1) return { success: false, message: "Sin tabla" };
      
      const sheetHeaders = values[headerRowIndex].map(h => String(h).toUpperCase().trim());
      const folioIndex = sheetHeaders.indexOf("FOLIO");
      const conceptoIndex = sheetHeaders.indexOf("CONCEPTO");

      const COL_MAP = {
        'ALTA': ['ALTA', 'CLIENTE', 'AREA'],
        'INVOLUCRADOS': ['INVOLUCRADOS', 'VENDEDOR', 'RESPONSABLE', 'REQUISITOR'],
        'RELOJ': ['RELOJ', 'DIAS'],
        'FECHA': ['FECHA', 'FECHA INICIO', 'FECHA DE INICIO'],
        'ESTATUS': ['ESTATUS', 'STATUS'],
        'PRIORIDAD': ['PRIORIDAD', 'PRIORIDADES'],
        'AVANCE': ['AVANCE', 'AVANCE %']
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
      for (let i = headerRowIndex + 1; i < values.length; i++) {
        if (values[i].join("|").toUpperCase().includes("TAREAS REALIZADAS")) break;
        
        const rowValFolio = folioIndex > -1 ? String(values[i][folioIndex]).trim().toUpperCase() : "";
        const rowValConcepto = conceptoIndex > -1 ? String(values[i][conceptoIndex]).trim().toUpperCase() : "";
        
        const inFolio = String(taskData['FOLIO']||"").trim().toUpperCase();
        const inConcepto = String(taskData['CONCEPTO']||"").trim().toUpperCase();

        if (inFolio && inFolio === rowValFolio) { targetRow = i + 1; break; }
        else if (inConcepto && inConcepto === rowValConcepto) { targetRow = i + 1; break; }
      }

      let isCompleted = false;
      const avanceKey = Object.keys(taskData).find(k => k.toUpperCase().includes("AVANCE"));
      if (avanceKey) {
          const val = String(taskData[avanceKey]).trim().replace('%', '');
          if (val === "100" || val === "1" || val === "1.0") isCompleted = true;
      }

      if (isCompleted) {
        let rowValues;
        if (targetRow !== -1) {
            rowValues = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).getValues()[0];
        } else {
            rowValues = new Array(sheet.getLastColumn()).fill("");
        }

        for (const key in taskData) {
           if (key.startsWith('_')) continue;
           const colIdx = getTargetColIdx(key);
           if (colIdx > -1) rowValues[colIdx] = taskData[key];
        }
        
        let historyIdx = -1;
        const allData = sheet.getDataRange().getValues();
        for(let r=0; r<allData.length; r++) {
            if(allData[r].join("|").toUpperCase().includes("TAREAS REALIZADAS")) { historyIdx = r+1; break; }
        }
        
        if (historyIdx === -1) {
           sheet.appendRow(["", "", "TAREAS REALIZADAS"]); 
           sheet.appendRow(values[headerRowIndex]); 
        }
        
        sheet.appendRow(rowValues);
        if (targetRow !== -1) { sheet.deleteRow(targetRow); }

        logSystemEvent(personName, "COMPLETE", `Archivado: ${taskData['CONCEPTO']}`);
        return { success: true, message: "Tarea completada y movida.", moved: true };
      }

      if (targetRow !== -1) {
        for (const key in taskData) {
          if (key.startsWith('_')) continue;
          const colIdx = getTargetColIdx(key);
          if (colIdx > -1) sheet.getRange(targetRow, colIdx + 1).setValue(taskData[key]);
        }
        logSystemEvent(personName, "UPDATE", `Editado: ${taskData['CONCEPTO']}`);
        return { success: true, message: "Guardado." };
      } else {
        const lastCol = sheet.getLastColumn();
        const headers = sheet.getRange(headerRowIndex+1, 1, 1, lastCol).getValues()[0].map(h => String(h).toUpperCase().trim());
        const newRow = headers.map(h => {
           const dataKey = Object.keys(taskData).find(k => {
              if (k.toUpperCase() === h) return true;
              for (const std in COL_MAP) {
                 if (COL_MAP[std].includes(k.toUpperCase()) && COL_MAP[std].includes(h)) return true;
              }
              return false;
           });
           return dataKey ? taskData[dataKey] : "";
        });
        
        sheet.insertRowAfter(headerRowIndex + 1);
        sheet.getRange(headerRowIndex + 2, 1, 1, newRow.length).setValues([newRow]);
        logSystemEvent(personName, "CREATE", `Nuevo: ${taskData['CONCEPTO']}`);
        return { success: true, message: "Creado." };
      }

    } catch(e) { return { success: false, message: e.toString() }; } 
    finally { lock.releaseLock(); }
  }
  return { success: false, message: "Ocupado." };
}

/* 7. DISTRIBUCIÓN MASIVA */
function apiSavePPCData(payload) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const items = Array.isArray(payload) ? payload : [payload];
      let sheet = findSheetSmart(APP_CONFIG.ppcSheetName);
      if (!sheet) { sheet = SS.insertSheet(APP_CONFIG.ppcSheetName); sheet.appendRow(["ID", "Especialidad", "Descripción", "Responsable", "Fecha", "Reloj", "Cumplimiento"]); }
      
      const fechaHoy = new Date();
      const fechaStr = Utilities.formatDate(fechaHoy, SS.getSpreadsheetTimeZone(), "dd/MM/yyyy");

      items.forEach(item => {
          const id = "PPC-" + Math.floor(Math.random() * 100000);
          sheet.appendRow([id, item.especialidad, item.concepto, item.responsable, fechaHoy, item.horas, item.cumplimiento, item.archivoUrl, item.comentarios, ""]);
          const responsables = String(item.responsable || "").split(",").map(s => s.trim()).filter(s => s);
          
          responsables.forEach(personName => {
            try {
              const taskData = {
                 'FOLIO': id,
                 'CONCEPTO': item.concepto,
                 'CLASIFICACION': item.clasificacion || "Media",
                 // SIN CLIENTE PARA EL PPC
                 'ALTA': item.especialidad,       
                 'INVOLUCRADOS': item.responsable,
                 'FECHA': fechaStr,
                 'RELOJ': item.horas,
                 'ESTATUS': "ASIGNADO",
                 'PRIORIDAD': item.prioridad,
                 'AVANCE': "0%"
              };
              apiUpdateTask(personName, taskData);
            } catch(err) { console.warn(err); }
          });
      });
      return { success: true, message: "Procesado." };
    } catch (e) { return { success: false, message: e.toString() }; } finally { lock.releaseLock(); }
  }
  return { success: false, message: "Ocupado." };
}

function uploadFileToDrive(d,t,n) { return {success:true, fileUrl:"http://mock.url"}; }
function apiFetchPPCData() { try { const s = findSheetSmart(APP_CONFIG.ppcSheetName); if(!s) return {success:true,data:[]}; const d = s.getRange(Math.max(2, s.getLastRow()-200),1,s.getLastRow(),10).getValues(); return {success:true, data: d.map(r=>({id:r[0],especialidad:r[1],concepto:r[2],responsable:r[3],fechaAlta:r[4],horas:r[5],cumplimiento:r[6]})).filter(x=>x.concepto).reverse()}; } catch(e){return {success:false}} }
function apiFetchSalesData() { try { const s = findSheetSmart(APP_CONFIG.salesSheetName); if(!s) return {success:true,data:[],headers:[]}; const d = s.getDataRange().getValues(); return {success:true, headers:d[0], data:d.slice(1).map(r=>{let o={};d[0].forEach((h,i)=>o[h]=r[i]);return o;})}; } catch(e){return {success:false}} }
function apiSaveSaleData(j) { const s = findSheetSmart(APP_CONFIG.salesSheetName); s.appendRow(s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(h=>j[h]||"")); return {success:true}; }
