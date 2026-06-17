// ============================================================
// QUINIELA FIFA MUNDIAL 2026 — Código.gs (Google Apps Script)
// ============================================================
// CONFIGURACIÓN: Antes de usar, ajusta estas constantes según
// tu entorno de Google Sheets y Drive.
// ============================================================

// ── ZONA HORARIA ─────────────────────────────────────────────
// Ajusta esto al identificador TZ de tu país. Ejemplos:
//   México:    "America/Mexico_City"
//   Argentina: "America/Argentina/Buenos_Aires"
//   Colombia:  "America/Bogota"
//   España:    "Europe/Madrid"
//   Chile:     "America/Santiago"
// Luego ve a: Archivo > Propiedades del proyecto > Zona horaria
// y selecciona la misma. Así el script y la UI coinciden.
const TIMEZONE = "America/Mexico_City";

// ── IDs DE GOOGLE SHEETS Y DRIVE ─────────────────────────────
// Abre tu Google Sheets → copia el ID de la URL:
// https://docs.google.com/spreadsheets/d/<<SPREADSHEET_ID>>/edit
const SPREADSHEET_ID = "PEGA_AQUI_EL_ID_DE_TU_SPREADSHEET";

// Carpeta de Drive donde se guardarán las fotos de los usuarios.
// Crea una carpeta en Drive, ábrela y copia el ID de la URL:
// https://drive.google.com/drive/folders/<<FOLDER_ID>>
const FOTO_FOLDER_ID = "PEGA_AQUI_EL_ID_DE_TU_CARPETA_EN_DRIVE";

// ── NOMBRES DE LAS HOJAS ──────────────────────────────────────
const HOJA_PARTIDOS   = "Partidos";
const HOJA_RESPUESTAS = "Respuestas";

// ── MINUTOS DE CORTE ANTES DEL PARTIDO ───────────────────────
// Si faltan menos de MINUTOS_CORTE para que inicie el partido,
// se marcará como BLOQUEADO y el usuario no podrá pronosticar.
const MINUTOS_CORTE = 5;

// ============================================================
// doGet — Punto de entrada de la Web App
// ============================================================
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Quiniela FIFA 2026")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// obtenerPartidosDeHoy
// Llamada desde el frontend con google.script.run
// Devuelve un objeto con:
//   { partidos: [...], fechaHoy: "YYYY-MM-DD", horaActual: "HH:MM" }
// Cada partido tiene: { id, equipoLocal, equipoVisitante, horaInicio,
//                        jornada, estado: "ACTIVO" | "BLOQUEADO" }
// ============================================================
function obtenerPartidosDeHoy() {
  try {
    const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja   = ss.getSheetByName(HOJA_PARTIDOS);
    const datos  = hoja.getDataRange().getValues(); // Incluye encabezados

    // Obtener fecha y hora actual en la zona horaria configurada
    const ahora      = new Date();
    const fechaHoy   = Utilities.formatDate(ahora, TIMEZONE, "yyyy-MM-dd");
    const horaActual = Utilities.formatDate(ahora, TIMEZONE, "HH:mm");

    const partidos = [];

    // La fila 0 es el encabezado; iteramos desde la fila 1
    // Estructura de columnas:
    // [0] ID_Partido | [1] Fecha (YYYY-MM-DD) | [2] Hora_Inicio (HH:MM)
    // [3] Equipo_Local | [4] Equipo_Visitante | [5] Jornada
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (!fila[0]) continue; // Salta filas vacías

      const idPartido      = String(fila[0]).trim();
      const fechaPartido   = Utilities.formatDate(new Date(fila[1]), TIMEZONE, "yyyy-MM-dd");
      const horaInicio     = String(fila[2]).trim(); // "HH:MM"
      const equipoLocal    = String(fila[3]).trim();
      const equipoVisitante = String(fila[4]).trim();
      const jornada        = String(fila[5]).trim();

      // Solo incluir partidos que sean HOY
      if (fechaPartido !== fechaHoy) continue;

      // Calcular si el partido ya está bloqueado
      const estado = calcularEstado(horaActual, horaInicio);

      partidos.push({
        id:              idPartido,
        equipoLocal:     equipoLocal,
        equipoVisitante: equipoVisitante,
        horaInicio:      horaInicio,
        jornada:         jornada,
        estado:          estado  // "ACTIVO" o "BLOQUEADO"
      });
    }

    return {
      exito:       true,
      partidos:    partidos,
      fechaHoy:    fechaHoy,
      horaActual:  horaActual
    };

  } catch (e) {
    Logger.log("Error en obtenerPartidosDeHoy: " + e.message);
    return { exito: false, error: e.message, partidos: [] };
  }
}

// ============================================================
// calcularEstado (función auxiliar interna)
// Compara la hora actual con la hora de inicio del partido.
// Si ya pasó o faltan menos de MINUTOS_CORTE → BLOQUEADO
// ============================================================
function calcularEstado(horaActualStr, horaInicioStr) {
  // Convierte "HH:MM" a minutos totales desde medianoche
  const toMinutos = (hhmm) => {
    const partes = hhmm.split(":");
    return parseInt(partes[0], 10) * 60 + parseInt(partes[1], 10);
  };

  const minutosActual = toMinutos(horaActualStr);
  const minutosInicio = toMinutos(horaInicioStr);

  // Si la hora actual es >= (inicio - margen), se bloquea
  if (minutosActual >= minutosInicio - MINUTOS_CORTE) {
    return "BLOQUEADO";
  }
  return "ACTIVO";
}

// ============================================================
// subirFoto
// Recibe el archivo en base64 desde el frontend y lo guarda
// en la carpeta de Drive configurada.
// Devuelve la URL pública del archivo subido.
// ============================================================
function subirFoto(base64Data, nombreArchivo, mimeType) {
  try {
    const folder   = DriveApp.getFolderById(FOTO_FOLDER_ID);
    const decoded  = Utilities.base64Decode(base64Data);
    const blob     = Utilities.newBlob(decoded, mimeType, nombreArchivo);
    const archivo  = folder.createFile(blob);

    // Hacer el archivo accesible con enlace
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { exito: true, url: archivo.getUrl() };
  } catch (e) {
    Logger.log("Error en subirFoto: " + e.message);
    return { exito: false, error: e.message, url: "" };
  }
}

// ============================================================
// guardarRespuestas
// Recibe del frontend:
//   { nombre, correo, urlFoto, predicciones: [{idPartido, local, visitante}] }
// Guarda una fila por predicción en la hoja "Respuestas".
// ============================================================
function guardarRespuestas(datos) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja  = ss.getSheetByName(HOJA_RESPUESTAS);
    const ts    = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss");

    // Validar que no haya predicciones de partidos bloqueados
    // (segunda capa de seguridad en el servidor)
    const ahora      = new Date();
    const horaActual = Utilities.formatDate(ahora, TIMEZONE, "HH:mm");

    // Obtener datos de partidos para verificar estados
    const hojaPartidos = ss.getSheetByName(HOJA_PARTIDOS);
    const datosPartidos = hojaPartidos.getDataRange().getValues();
    const mapaPartidos  = {};

    for (let i = 1; i < datosPartidos.length; i++) {
      const fila = datosPartidos[i];
      if (!fila[0]) continue;
      const idP = String(fila[0]).trim();
      mapaPartidos[idP] = { horaInicio: String(fila[2]).trim() };
    }

    const filasAGuardar = [];

    for (const pred of datos.predicciones) {
      const idPartido = String(pred.idPartido);

      // Verificar estado en el servidor (anti-trampa)
      if (mapaPartidos[idPartido]) {
        const estado = calcularEstado(horaActual, mapaPartidos[idPartido].horaInicio);
        if (estado === "BLOQUEADO") {
          // Ignorar silenciosamente predicciones de partidos bloqueados
          continue;
        }
      }

      // Estructura: Timestamp | Nombre | Correo | URL_Foto | ID_Partido
      //             | Prediccion_Local | Prediccion_Visitante
      filasAGuardar.push([
        ts,
        datos.nombre,
        datos.correo,
        datos.urlFoto || "",
        idPartido,
        parseInt(pred.local, 10),
        parseInt(pred.visitante, 10)
      ]);
    }

    if (filasAGuardar.length > 0) {
      hoja.getRange(
        hoja.getLastRow() + 1,
        1,
        filasAGuardar.length,
        7
      ).setValues(filasAGuardar);
    }

    return { exito: true, guardados: filasAGuardar.length };

  } catch (e) {
    Logger.log("Error en guardarRespuestas: " + e.message);
    return { exito: false, error: e.message };
  }
}

// ============================================================
// NOTAS DE CONFIGURACIÓN DEL SPREADSHEET
// ============================================================
//
// HOJA "Partidos" — Encabezados exactos en la FILA 1:
// ┌────────────┬────────────┬─────────────┬──────────────┬─────────────────┬─────────┐
// │ ID_Partido │   Fecha    │ Hora_Inicio │ Equipo_Local │ Equipo_Visitante │ Jornada │
// ├────────────┼────────────┼─────────────┼──────────────┼─────────────────┼─────────┤
// │   P001     │ 2026-06-11 │    18:00    │   México     │     Polonia      │    2    │
// │   P002     │ 2026-06-11 │    21:00    │   Brasil     │    Argentina     │    2    │
// └────────────┴────────────┴─────────────┴──────────────┴─────────────────┴─────────┘
//
// IMPORTANTE: La columna "Fecha" debe estar formateada como TEXTO plano
// (no como fecha de Sheets) para evitar desfases de zona horaria al leer.
// Selecciona la columna B → Formato → Número → Texto plano.
//
// HOJA "Respuestas" — Encabezados exactos en la FILA 1:
// Timestamp | Nombre | Correo | URL_Foto | ID_Partido |
// Prediccion_Local | Prediccion_Visitante
//
// PUBLICAR COMO WEB APP:
// 1. Implementar > Nueva implementación > Web App
// 2. Ejecutar como: "Yo (tu correo)"
// 3. Quién tiene acceso: "Cualquier persona"
// 4. Haz clic en "Implementar" y copia la URL generada.
// ============================================================
