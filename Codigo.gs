/**
 * ============================================================================
 * 🚀 SCRIPT MASTER: HOLTMONT ONE ENGINE V20 (ARQUITECTURA MONOLÍTICA FINAL)
 * ============================================================================
 * * CARACTERÍSTICAS:
 * 1. Conexión ÚNICA al archivo maestro (ID 16K695...).
 * 2. Detector Inteligente: Salta los cuadros de resumen (filas 1-13) y encuentra los datos reales.
 * 3. Lista de Personal completa y desplegada.
 */

// --- 1. CONFIGURACIÓN GLOBAL ---
const CONFIG = {
  // ✅ ID DE TU ARCHIVO MAESTRO (Donde unificaste todas las pestañas)
  // Este es el ID real extraído de tu link.
  SPREADSHEET_ID: "16K695zTHfZd4KS9JHcxdfvfJnCUwb2rfqTN3k_TFzB0", 

  // ID de Ventas (Si Antonia sigue usando su hoja aparte, déjalo así. Si ya la moviste, el script la buscará en el maestro igual)
  SALES_SHEET_ID: "1uyO_6Wwz-VlCeX9LyiL-pZG9MVfP0PDIjK-58TMxSPM",

  // Carpetas de Drive (Tus IDs originales)
  FOLDERS: {
    'doc': '1466-Gl1YnkU8rnKgZVFrOYKdik6PRtKc', 
    'planos': '11QQ4LG6SEKzgfUw_d0QivLt1yHM-dJSw',
    'foto': '1D7UdqE1VXbF2nXFHgPoU9DmzZTE2lmZ5', 
    'corr': '1WOu1YhGXGEGK4uroyMs6VOQZ2tk8FZQr',
    'rep': '1hPtye2D_9BbuFVverhCZBi2KgkurwVC'
  },
  
  // Usuarios y Roles
  USERS: {
    'admin123': { role: 'ADMIN', name: 'Super Admin' },
    'ventas2025': { role: 'SALES', name: 'Antonia (Ventas)' }
  }
};

// --- 2. INICIALIZACIÓN ---
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Holtmont Workspace V20')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function apiLogin(password) {
  const user = CONFIG.USERS[password];
  if (user) return { success: true, role: user.role, name: user.name };
  return { success: false, message: 'Contraseña incorrecta.' };
}

// --- 3. CONFIGURACIÓN DEL PERSONAL (LISTA COMPLETA) ---
function getSystemConfig(role) {
  
  // Esta lista es VITAL. El nombre aquí debe coincidir con el nombre de la pestaña en el Excel.
  const fullStaff = [
      // HVAC
      { name: 'EDUARDO MANZANARES', dept: 'HVAC' },
      { name: 'JUAN JOSE SANCHEZ', dept: 'HVAC' },
      { name: 'SELENE BALDONADO', dept: 'HVAC' },
      { name: 'ROLANDO MORENO', dept: 'HVAC' },
      
      // ELECTROMECÁNICA
      { name: 'INGE GALLARDO', dept: 'ELECTRO' },
      { name: 'SEBASTIAN PADILLA', dept: 'ELECTRO' },
      { name: 'JEHU MARTINEZ', dept: 'ELECTRO' },
      { name: 'MIGUEL GONZALEZ', dept: 'ELECTRO' },
      { name: 'ALICIA RIVERA', dept: 'ELECTRO' },
      
      // CONSTRUCCIÓN
      { name: 'RICARDO MENDO', dept: 'CONST' },
      { name: 'CARLOS MENDEZ', dept: 'CONST' },
      { name: 'REYNALDO GARCIA', dept: 'CONST' },
      { name: 'INGE OLIVO', dept: 'CONST' },
      { name: 'EDUARDO TERAN', dept: 'CONST' },
      { name: 'EDGAR HOLT', dept: 'CONST' },
      { name: 'ALEXIS TORRES', dept: 'CONST' },
      { name: 'TERESA GARZA', dept: 'CONST' },
      { name: 'RAMIRO RODRIGUEZ', dept: 'CONST' },
      { name: 'GUILLERMO DAMICO', dept: 'CONST' },
      { name: 'RUBEN PESQUEDA', dept: 'CONST' },
      
      // COMPRAS
      { name: 'JUDITH ECHAVARRIA', dept: 'COMPRAS' },
      { name: 'GISELA DOMINGUEZ', dept: 'COMPRAS' },
      { name: 'VANESSA DE LARA', dept: 'COMPRAS' },
      { name: 'NELSON MALDONADO', dept: 'COMPRAS' },
      { name: 'VICTOR ALMACEN', dept: 'COMPRAS' },
      
      // EHS
      { name: 'DIMAS RAMOS', dept: 'EHS' },
      { name: 'CITLALI GOMEZ', dept: 'EHS' },
      { name: 'AIMEE RAMIREZ', dept: 'EHS' },
      
      // MAQUINARIA (Nota: Edgar y Alexis aparecen duplicados intencionalmente si tienen doble función)
      { name: 'EDGAR HOLT', dept: 'MAQ' },
      { name: 'ALEXIS TORRES', dept: 'MAQ' },
      
      // DISEÑO
      { name: 'ANGEL SALINAS', dept: 'DISENO' },
      { name: 'EDGAR LOPEZ', dept: 'DISENO' },
      
      // ADMINISTRACIÓN
      { name: 'ANTONIO CABRERA', dept: 'ADMIN' },
      { name: 'EDUARDO BENITEZ', dept: 'ADMIN' },
      { name: 'DANIELA CASTRO', dept: 'ADMIN' },
      { name: 'LILIANA MARTINEZ', dept: 'ADMIN' },
      { name: 'LAURA HUERTA', dept: 'ADMIN' },
      { name: 'ANTONIA PINEDA', dept: 'ADMIN' },
      { name: 'JUANY RODRIGUEZ', dept: 'ADMIN' },
      { name: 'DANIA GONZALEZ', dept: 'ADMIN' },
      { name: 'ROCIO CASTRO', dept: 'ADMIN' },
      { name: 'LUIS CARLOS', dept: 'ADMIN' },
      { name: 'ANTONIO SALAZAR', dept: 'ADMIN' },
      
      // VENTAS
      { name: 'ALFONSO CORREA', dept: 'VENTAS' },
      { name: 'CESAR GOMEZ', dept: 'VENTAS' }
  ];

  // Configuración para ANTONIA
  if (role === 'SALES') {
    return {
      role: role,
      departments: {},
      staff: [],
      specialModules: [
        { id: 'SALES_MASTER', label: 'Control Ventas', icon: 'fa-chart-line', color: '#EA4335', type: 'sales_native' }
      ]
    };
  }

  // Configuración para ADMIN
  if (role === 'ADMIN') {
    return {
      role: role,
      departments: {
        'ADMIN': { label: 'Administración', color: '#4285F4', icon: 'fa-desktop' },
        'HVAC': { label: 'HVAC', color: '#5C6BC0', icon: 'fa-fan' },
        'ELECTRO': { label: 'Electromecánica', color: '#26A69A', icon: 'fa-bolt' },
        'CONST': { label: 'Construcción', color: '#FFA000', icon: 'fa-hard-hat' },
        'EHS': { label: 'EHS', color: '#4CAF50', icon: 'fa-shield-alt' },
        'MAQ': { label: 'Maquinaria', color: '#607D8B', icon: 'fa-tractor' },
        'DISENO': { label: 'Diseño', color: '#9C27B0', icon: 'fa-tools' },
        'COMPRAS': { label: 'Compras', color: '#795548', icon: 'fa-warehouse' },
        'VENTAS': { label: 'Ventas', color: '#198754', icon: 'fa-chart-line' }
      },
      specialModules: [
        { id: 'PPC_MASTER', label: 'Panel PPC', icon: 'fa-tasks', color: '#2c5aa0', type: 'ppc_native', url: '' },
        { id: 'SALES_MASTER', label: 'Supervisión Ventas', icon: 'fa-store', color: '#EA4335', type: 'sales_native' }, 
        
        // Tus pestañas de coordinadores (Links externos originales se mantienen si los necesitas)
        { 
          id: 'COOR', label: 'Coordinadores', icon: 'fa-users-cog', color: '#343a40', type: 'tabs',
          tabs: [
            { name: 'Construcción', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQnaeRzEIyd99Cjbq2CT2gETZSIWd-BSH_wyZ8LQyqQOmyv6x76pjOrXG6z05xzVk7W3ULkvPMJDum0/pubhtml?gid=458557491&single=true' },
            { name: 'HVAC', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQnaeRzEIyd99Cjbq2CT2gETZSIWd-BSH_wyZ8LQyqQOmyv6x76pjOrXG6z05xzVk7W3ULkvPMJDum0/pubhtml?gid=1698456178&single=true' },
            { name: 'Electromecánica', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQnaeRzEIyd99Cjbq2CT2gETZSIWd-BSH_wyZ8LQyqQOmyv6x76pjOrXG6z05xzVk7W3ULkvPMJDum0/pubhtml?gid=76398777&single=true' }
          ]
        },
        { 
          id: 'AREAS', label: 'Demás Áreas', icon: 'fa-layer-group', color: '#6610f2', type: 'tabs',
          tabs: [
            { name: 'Judith E.', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQnaeRzEIyd99Cjbq2CT2gETZSIWd-BSH_wyZ8LQyqQOmyv6x76pjOrXG6z05xzVk7W3ULkvPMJDum0/pubhtml?gid=1384864037&single=true' },
            { name: 'Ramiro R.', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQnaeRzEIyd99Cjbq2CT2gETZSIWd-BSH_wyZ8LQyqQOmyv6x76pjOrXG6z05xzVk7W3ULkvPMJDum0/pubhtml?gid=482040858&single=true' },
            { name: 'Alfonso C.', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQnaeRzEIyd99Cjbq2CT2gETZSIWd-BSH_wyZ8LQyqQOmyv6x76pjOrXG6z05xzVk7W3ULkvPMJDum0/pubhtml?gid=2109402945&single=true' },
            { name: 'Teresa G.', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQnaeRzEIyd99Cjbq2CT2gETZSIWd-BSH_wyZ8LQyqQOmyv6x76pjOrXG6z05xzVk7W3ULkvPMJDum0/pubhtml?gid=1751645258&single=true' },
            { name: 'Guillermo D.', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQnaeRzEIyd99Cjbq2CT2gETZSIWd-BSH_wyZ8LQyqQOmyv6x76pjOrXG6z05xzVk7W3ULkvPMJDum0/pubhtml?gid=218034731&single=true' }
          ]
        }
      ],
      staff: fullStaff
    };
  }
}

// --- 4. MOTOR DE DETECCIÓN INTELIGENTE (ESTE ES EL SECRETO) ---

/**
 * Busca una pestaña por nombre, ignorando mayúsculas, acentos y espacios extra.
 */
function findSheetFlexible(ss, targetName) {
  if (!targetName) return null;
  const cleanTarget = targetName.toString().toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, '');
  
  const sheets = ss.getSheets();
  for (let sheet of sheets) {
    const cleanSheetName = sheet.getName().toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, '');
    if (cleanTarget === cleanSheetName) return sheet;
  }
  return null;
}

/**
 * ⚠️ LA SOLUCIÓN: ESCANER DE ENCABEZADOS
 * Recorre las primeras 50 filas buscando dónde empieza REALMENTE la tabla.
 * Busca las palabras clave: "ID", "Especialidad" y "Descripción".
 */
function findHeaderRowIndex(values) {
  for (let i = 0; i < Math.min(values.length, 50); i++) {
    const rowStr = values[i].join(' ').toLowerCase();
    // Validamos que la fila tenga estas 3 columnas clave para confirmar que es la cabecera
    if (rowStr.includes('id') && (rowStr.includes('especialidad') || rowStr.includes('descripcion') || rowStr.includes('actividad'))) {
      return i; // Retorna el índice de la fila correcta
    }
  }
  return 0; // Si no encuentra nada, usa la primera fila por defecto
}

// --- 5. LÓGICA DE LECTURA Y ESCRITURA ---

/**
 * Lee los datos de cualquier hoja detectando automáticamente su estructura.
 */
function readDataFromSheet(sheet) {
  try {
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // 1. Detectar dónde empiezan los datos
    const headerRowIndex = findHeaderRowIndex(values);
    
    // 2. Extraer encabezados y filas de datos
    const headers = values[headerRowIndex].map(h => h.toString().toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""));
    const dataRows = values.slice(headerRowIndex + 1);

    // 3. Mapeo Dinámico (Busca en qué columna cayó cada dato)
    const colMap = {
      id: headers.findIndex(h => h.includes('id')),
      especialidad: headers.findIndex(h => h.includes('especialidad')),
      concepto: headers.findIndex(h => h.includes('descripcion') || h.includes('actividad') || h.includes('concepto')),
      responsable: headers.findIndex(h => h.includes('responsable')),
      fechaAlta: headers.findIndex(h => h.includes('fecha') && (h.includes('alta') || h.includes('compromiso'))),
      horas: headers.findIndex(h => h.includes('reloj') || h.includes('tiempo') || h.includes('transcurrido')),
      prioridad: headers.findIndex(h => h.includes('prioridad')),
      cumplimiento: headers.findIndex(h => h.includes('cumplimiento')),
      archivoUrl: headers.findIndex(h => h.includes('archivo')),
      comentarios: headers.findIndex(h => h.includes('comentario') && (h.includes('curso') || h.includes('actual') || h.includes('semana'))),
      comentariosPrevios: headers.findIndex(h => h.includes('comentario') && (h.includes('previa') || h.includes('anterior'))),
      horaFin: headers.findIndex(h => h.includes('fin'))
    };

    // 4. Construir objetos de datos limpios
    return dataRows.map(row => {
      const getVal = (idx) => (idx >= 0 && row[idx] !== undefined) ? row[idx] : '';
      
      let fechaVal = getVal(colMap.fechaAlta);
      if (fechaVal instanceof Date) {
        fechaVal = Utilities.formatDate(fechaVal, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }

      return {
        id: getVal(colMap.id).toString(),
        especialidad: getVal(colMap.especialidad).toString(),
        concepto: getVal(colMap.concepto).toString(),
        responsable: getVal(colMap.responsable).toString(),
        fechaAlta: fechaVal,
        horas: getVal(colMap.horas).toString(),
        prioridad: getVal(colMap.prioridad).toString(),
        cumplimiento: getVal(colMap.cumplimiento).toString(),
        archivoUrl: getVal(colMap.archivoUrl).toString(),
        comentarios: getVal(colMap.comentarios).toString(),
        comentariosPrevios: getVal(colMap.comentariosPrevios).toString(),
        horaFin: getVal(colMap.horaFin).toString()
      };
    })
    .filter(item => item.id && item.id.toString().trim() !== '') // Eliminar filas vacías
    .reverse();

  } catch (e) {
    return []; 
  }
}

/**
 * ⚡ API: TRACKER INDIVIDUAL
 * Busca la pestaña del empleado en el archivo maestro.
 */
function apiFetchStaffTrackerData(staffName) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = findSheetFlexible(ss, staffName);
    
    if (!sheet) return { success: false, message: `No se encontró la pestaña "${staffName}".` };
    
    const data = readDataFromSheet(sheet);
    return { success: true, data: data };
  } catch (e) {
    return { success: false, message: "Error Tracker: " + e.message };
  }
}

/**
 * ⚡ API: PANEL GENERAL (PPC)
 * Busca la primera hoja válida del archivo maestro.
 */
function apiFetchPPCData() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    // Busca la primera hoja que tenga estructura de tabla (datos)
    let sheet = ss.getSheets()[0];
    for (let s of ss.getSheets()) {
       const vals = s.getRange(1,1,30,20).getValues();
       if (findHeaderRowIndex(vals) > 0) { // Si encuentra cabeceras válidas
         sheet = s; 
         break;
       }
    }
    const data = readDataFromSheet(sheet);
    return { success: true, data: data };
  } catch(e) {
    return { success: false, message: "Error Panel PPC: " + e.message };
  }
}

/**
 * ⚡ API: GUARDAR DATOS (PPC)
 */
function apiSavePPCData(formData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // Intenta guardar en hoja "Data" o la primera que encuentre
    let sheet = findSheetFlexible(ss, "Data");
    if (!sheet) sheet = ss.getSheets()[0];
    
    const now = new Date();
    const fechaHoy = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    // Mapeo simple para guardado
    const newRow = [
      `TK-${Math.floor(Math.random() * 10000)}`, // ID
      formData.especialidad, 
      formData.concepto, 
      formData.responsable, 
      fechaHoy, 
      formData.horas, 
      formData.prioridad || "Media",
      formData.cumplimiento || "NO", 
      formData.archivoUrl || "", 
      formData.comentarios, 
      "" 
    ];
    
    sheet.appendRow(newRow);
    return { success: true, message: "Guardado exitosamente." };
  } catch (e) {
    return { success: false, message: "Error al guardar: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

// --- UTILS: SUBIDA DE ARCHIVOS ---
function uploadFileToDrive(data, type, name) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(data.substr(data.indexOf('base64,')+7)), data.substring(5, data.indexOf(';')), name);
    const folderId = CONFIG.FOLDERS[type] || CONFIG.FOLDERS['doc'];
    const file = DriveApp.getFolderById(folderId).createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, fileUrl: file.getUrl(), folder: type };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// --- MÓDULO VENTAS (ANTONIA) ---
function apiFetchSalesData() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // Busca pestaña de Ventas en el maestro
    let sheet = findSheetFlexible(ss, "Ventas");
    if (!sheet) sheet = findSheetFlexible(ss, "Sales");
    if (!sheet) {
        // Fallback: Intenta el ID externo antiguo si no está en el maestro
        try {
           const ssExt = SpreadsheetApp.openById(CONFIG.SALES_SHEET_ID);
           sheet = ssExt.getSheets()[0];
        } catch(e) { return { success: false, message: "No se encontró hoja de Ventas." }; }
    }

    const data = readDataFromSheet(sheet);
    
    // Si readData falla por formato, intento simple
    if (data.length === 0) {
       const values = sheet.getDataRange().getValues();
       const headers = values[0];
       const simpleData = values.slice(1).map(row => {
         const obj = {}; headers.forEach((h, i) => obj[h] = row[i]); return obj;
       });
       return { success: true, data: simpleData.reverse(), headers: headers };
    }

    return { success: true, data: data, headers: Object.keys(data[0] || {}) };

  } catch (e) {
    return { success: false, message: "Error Ventas: " + e.message };
  }
}

function apiSaveSaleData(formData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = findSheetFlexible(ss, "Ventas");
    if (!sheet) { // Fallback externo
        const ssExt = SpreadsheetApp.openById(CONFIG.SALES_SHEET_ID);
        sheet = ssExt.getSheets()[0];
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = headers.map(h => formData[h] || "");
    sheet.appendRow(newRow);
    return { success: true, message: "Venta registrada." };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}
