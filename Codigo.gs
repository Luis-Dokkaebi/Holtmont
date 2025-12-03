/**
 * ============================================================================
 * 🚀 SCRIPT MASTER: HOLTMONT ONE ENGINE (FINAL V8 - FULL FEATURES)
 * ============================================================================
 */

// --- CONFIGURACIÓN GLOBAL ---
const CONFIG = {
  PPC_SHEET_ID: "1iwlanqnhHGjo9MdZtbsymXu1zCsdjRONdxqfYV6icck",
  
  FOLDERS: {
    'doc': '1466-Gl1YnkU8rnKgZVFrOYKdik6PRtKc', 
    'planos': '11QQ4LG6SEKzgfUw_d0QivLt1yHM-dJSw',
    'foto': '1D7UdqE1VXbF2nXFHgPoU9DmzZTE2lmZ5', 
    'corr': '1WOu1YhGXGEGK4uroyMs6VOQZ2tk8FZQr',
    'rep': '1hPtye2D_9BbuFVverhCZBi2KgkurwVC'
  }
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Holtmont Workspace v8.0')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSystemConfig() {
  return {
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
      { id: 'ANTONIA', label: 'Antonia (Ventas)', icon: 'fa-store', color: '#EA4335', type: 'iframe', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQnaeRzEIyd99Cjbq2CT2gETZSIWd-BSH_wyZ8LQyqQOmyv6x76pjOrXG6z05xzVk7W3ULkvPMJDum0/pubhtml?gid=0&single=true' },
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
          { name: 'Guillermo D.', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQnaeRzEIyd99Cjbq2CT2gETZSIWd-BSH_wyZ8LQyqQOmyv6x76pjOrXG6z05xzVk7W3ULkvPMJDum0/pubhtml?gid=218034731&single=true' },
          { name: 'Angel S.', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQnaeRzEIyd99Cjbq2CT2gETZSIWd-BSH_wyZ8LQyqQOmyv6x76pjOrXG6z05xzVk7W3ULkvPMJDum0/pubhtml?gid=1056760648&single=true' },
          { name: 'Sebastian P.', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQnaeRzEIyd99Cjbq2CT2gETZSIWd-BSH_wyZ8LQyqQOmyv6x76pjOrXG6z05xzVk7W3ULkvPMJDum0/pubhtml?gid=1336557549&single=true' },
          { name: 'Eduardo M.', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQnaeRzEIyd99Cjbq2CT2gETZSIWd-BSH_wyZ8LQyqQOmyv6x76pjOrXG6z05xzVk7W3ULkvPMJDum0/pubhtml?gid=2145395127&single=true' },
          { name: 'Juan Jose S.', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQnaeRzEIyd99Cjbq2CT2gETZSIWd-BSH_wyZ8LQyqQOmyv6x76pjOrXG6z05xzVk7W3ULkvPMJDum0/pubhtml?gid=1327157580&single=true' },
          { name: 'Cesar G.', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQnaeRzEIyd99Cjbq2CT2gETZSIWd-BSH_wyZ8LQyqQOmyv6x76pjOrXG6z05xzVk7W3ULkvPMJDum0/pubhtml?gid=1708938131&single=true' }
        ]
      },
      { id: 'PRUEBA', label: 'Angel Salinas (Prueba)', icon: 'fa-vial', color: '#FFD700', type: 'iframe', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTbT1fB4HpItMHpSf6J9_4zQ0cfLYg72kB0mTSITeKlMLt7IIvrQ_6H7u-a9orA_teEbfUeK1mftwMs/pubhtml?gid=267375035&single=true' }
    ],
    
    // LISTA COMPLETA DE PERSONAL
    staff: [
      { name: 'EDUARDO MANZANARES', dept: 'HVAC', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=853670968&single=true' },
      { name: 'JUAN JOSE SANCHEZ', dept: 'HVAC', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=376327929&single=true' },
      { name: 'SELENE BALDONADO', dept: 'HVAC', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=2059318779&single=true' },
      { name: 'ROLANDO MORENO', dept: 'HVAC', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=763016000&single=true' },
      { name: 'INGE GALLARDO', dept: 'ELECTRO', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=729505665&single=true' },
      { name: 'SEBASTIAN PADILLA', dept: 'ELECTRO', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=2059550884&single=true' },
      { name: 'JEHU MARTINEZ', dept: 'ELECTRO', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=2010526863&single=true' },
      { name: 'MIGUEL GONZALEZ', dept: 'ELECTRO', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=382932794&single=true' },
      { name: 'ALICIA RIVERA', dept: 'ELECTRO', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=916323918&single=true' },
      { name: 'RICARDO MENDO', dept: 'CONST', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vT6n0wM8NTcpNHqORRiVAmZx0PL8OJ-Z1pMEWPkMFIjg_KSvZj5obf0GHa6Hs8mF85HIbFpDVw7ho2X/pubhtml?gid=190530424&single=true' },
      { name: 'CARLOS MENDEZ', dept: 'CONST', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=1913258556&single=true' },
      { name: 'REYNALDO GARCIA', dept: 'CONST', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=1149435901&single=true' },
      { name: 'INGE OLIVO', dept: 'CONST', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=16110288&single=true' },
      { name: 'EDUARDO TERAN', dept: 'CONST', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=308559493&single=true' },
      { name: 'EDGAR HOLT', dept: 'CONST', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=1377687670&single=true' },
      { name: 'ALEXIS TORRES', dept: 'CONST', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=354167442&single=true' },
      { name: 'TERESA GARZA', dept: 'CONST', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=575319228&single=true' },
      { name: 'RAMIRO RODRIGUEZ', dept: 'CONST', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=917098163&single=true' },
      { name: 'GUILLERMO DAMICO', dept: 'CONST', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=1650054684&single=true' },
      { name: 'RUBEN PESQUEDA', dept: 'CONST', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=97035860&single=true' },
      { name: 'JUDITH ECHAVARRIA', dept: 'COMPRAS', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=1400855940&single=true' },
      { name: 'GISELA DOMINGUEZ', dept: 'COMPRAS', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=1945398995&single=true' },
      { name: 'VANESSA DE LARA', dept: 'COMPRAS', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=945522475&single=true' },
      { name: 'NELSON MALDONADO', dept: 'COMPRAS', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=1522744000&single=true' },
      { name: 'VICTOR ALMACEN', dept: 'COMPRAS', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=890766536&single=true' },
      { name: 'DIMAS RAMOS', dept: 'EHS', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=1682107802&single=true' },
      { name: 'CITLALI GOMEZ', dept: 'EHS', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=531253869&single=true' },
      { name: 'AIMEE RAMIREZ', dept: 'EHS', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=1607186101&single=true' },
      { name: 'EDGAR HOLT', dept: 'MAQ', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vT6n0wM8NTcpNHqORRiVAmZx0PL8OJ-Z1pMEWPkMFIjg_KSvZj5obf0GHa6Hs8mF85HIbFpDVw7ho2X/pubhtml?gid=2008532628&single=true' },
      { name: 'ALEXIS TORRES', dept: 'MAQ', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vT6n0wM8NTcpNHqORRiVAmZx0PL8OJ-Z1pMEWPkMFIjg_KSvZj5obf0GHa6Hs8mF85HIbFpDVw7ho2X/pubhtml?gid=549117226&single=true' },
      { name: 'ANGEL SALINAS', dept: 'DISENO', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=722740650&single=true' },
      { name: 'EDGAR LOPEZ', dept: 'DISENO', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQgjqB9tMWvD22iat72wn-H70yIXONPjXdpfT_2sghFYDAu-JTFuG-oU5yOihDGkEiYFZy6gfNe72UT/pubhtml?gid=1771015939&single=true' },
      { name: 'ANTONIO CABRERA', dept: 'ADMIN', url: '#' },
      { name: 'EDUARDO BENITEZ', dept: 'ADMIN', url: '#' },
      { name: 'DANIELA CASTRO', dept: 'ADMIN', url: '#' },
      { name: 'LILIANA MARTINEZ', dept: 'ADMIN', url: '#' },
      { name: 'LAURA HUERTA', dept: 'ADMIN', url: '#' },
      { name: 'ANTONIA PINEDA', dept: 'ADMIN', url: '#' },
      { name: 'JUANY RODRIGUEZ', dept: 'ADMIN', url: '#' },
      { name: 'DANIA GONZALEZ', dept: 'ADMIN', url: '#' },
      { name: 'ROCIO CASTRO', dept: 'ADMIN', url: '#' },
      { name: 'LUIS CARLOS', dept: 'ADMIN', url: '#' },
      { name: 'ANTONIO SALAZAR', dept: 'ADMIN', url: '#' },
      { name: 'ALFONSO CORREA', dept: 'VENTAS', url: '#' },
      { name: 'CESAR GOMEZ', dept: 'VENTAS', url: '#' }
    ]
  };
}

/**
 * ⚡ API: GUARDAR DATOS PPC
 */
function apiSavePPCData(formData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(CONFIG.PPC_SHEET_ID);
    
    let sheet = ss.getSheetByName("Data");
    if (!sheet) sheet = ss.getSheets()[0];

    // Genera la fecha y hora actual con formato estricto (dd/MM/yyyy HH:mm)
    const now = new Date();
    const fechaHoy = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    
    // ORDEN FINAL DE COLUMNAS (Coincide con A-N)
    const newRow = [
      // Columna A (0)
      `TK-${Math.floor(Math.random() * 10000)}`, 
      // Columna B (1)
      formData.especialidad,
      // Columna C (2)
      formData.concepto,
      // Columna D (3)
      formData.responsable, 
      // Columna E (4) <-- Usa la fecha y hora CORRECTA
      fechaHoy,
      // Columna F (5)
      formData.horas,
      // Columna G (6) <-- Prioridad: Urgente, Media, Baja
      formData.prioridad || "Media",
      // Columna H (7)
      formData.cumplimiento || "NO",
      // Columna I (8)
      formData.archivoUrl || "",
      // Columna J (9)
      formData.comentarios,
      // Columna K (10)
      formData.comentariosPrevios || "", 
      // Columna L (11)
      formData.clasificacion,
      // Columna M (12) <-- Riesgos: Bajo, Medio, Alto, Catastrófico
      formData.riesgos,       
      // Columna N (13)
      formData.horaFin        
    ];

    sheet.appendRow(newRow);
    return { success: true, message: "Actividad registrada exitosamente." };
    
  } catch (e) {
    return { success: false, message: "Error al guardar: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ⚡ API: OBTENER TODOS LOS DATOS PPC (Optimizado para Listado General)
 */
function apiFetchPPCData() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.PPC_SHEET_ID);
    const sheet = ss.getSheetByName("Data");
    if (!sheet) throw new Error("La hoja 'Data' no se encontró. Revise el nombre.");

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    if (values.length <= 1) { 
      return { success: true, data: [] };
    }

    const headers = [
      'id', 'especialidad', 'concepto', 'responsable', 'fechaAlta', 'horas', 
      'prioridad', 'cumplimiento', 'archivoUrl', 'comentarios', 'comentariosPrevios', 
      'clasificacion', 'riesgos', 'horaFin'
    ];
    
    const dataRows = values.slice(1); 
    const numColumnsToUse = Math.min(values[0].length, headers.length);

    const structuredData = dataRows
      .map(row => {
        const obj = {};
        for (let i = 0; i < numColumnsToUse; i++) {
          obj[headers[i]] = row[i];
        }
        return obj;
      })
      .filter(row => row.id && row.id.toString().startsWith('TK-'));

    return { success: true, data: structuredData };

  } catch (e) {
    Logger.log(`Error en apiFetchPPCData: ${e.message}`);
    return { success: false, message: `Error al cargar datos PPC: ${e.message}` };
  }
}

/**
 * ⚡ API: OBTENER TAREAS DEL EMPLEADO (Optimizado para Tracker Nativo)
 */
function apiFetchStaffTrackerData(staffName) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.PPC_SHEET_ID);
    const sheet = ss.getSheetByName("Data");
    if (!sheet) throw new Error("La hoja 'Data' no se encontró.");

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    if (values.length <= 1) {
      return { success: true, data: [] };
    }

    const headers = [
      'id', 'especialidad', 'concepto', 'responsable', 'fechaAlta', 'horas', 
      'prioridad', 'cumplimiento', 'archivoUrl', 'comentarios', 'comentariosPrevios', 
      'clasificacion', 'riesgos', 'horaFin'
    ];
    
    const dataRows = values.slice(1); 
    const numColumnsToUse = Math.min(values[0].length, headers.length);
    const normalizedStaffName = staffName.toUpperCase().trim();

    const structuredData = dataRows
      .filter(row => {
        const responsableCell = row[3] ? row[3].toString().toUpperCase() : ''; 
        return responsableCell.includes(normalizedStaffName);
      })
      .map(row => {
        const obj = {};
        for (let i = 0; i < numColumnsToUse; i++) {
          let value = row[i];
          const header = headers[i];
          
          // Manejo y normalización de valores
          if (value === null || value === undefined) {
              value = '';
          } else if (typeof value === 'object' && value instanceof Date) {
              
              if (header === 'comentariosPrevios' || header === 'comentarios') {
                  value = ''; 
              } else if (header === 'fechaAlta') {
                  
                  // ✅ SOLUCIÓN FINAL: Blindaje total. Si la fecha es inválida o la base (1899), la eliminamos.
                  // Si no es la fecha base, mostramos SOLO la hora (HH:mm).
                  if (value.getTime() <= 86400000 || value.getFullYear() < 1970) { 
                      value = ''; 
                  } else {
                      value = Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
                  }
              }
          } else if (typeof value === 'number') {
              // Si es un número (serial) y es 0, lo tratamos como vacío.
              value = value === 0 ? '' : value.toString();
          }

          obj[header] = value.toString().trim();
        }
        return obj;
      })
      .reverse();

    return { success: true, data: structuredData };

  } catch (e) {
    Logger.log(`Error en apiFetchStaffTrackerData: ${e.message}`);
    return { success: false, message: `Error al cargar datos del empleado: ${e.message}` };
  }
}


/**
 * ⚡ API: SUBIR ARCHIVOS
 */
function uploadFileToDrive(data, type, name) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(data.substr(data.indexOf('base64,')+7)), data.substring(5, data.indexOf(';')), name);
    const folderId = CONFIG.FOLDERS[type] || CONFIG.FOLDERS['doc'];
    const file = DriveApp.getFolderById(folderId).createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { success: true, fileUrl: file.getUrl(), folder: type };
  } catch (e) { 
    return { success: false, message: e.toString() }; 
  }
}
