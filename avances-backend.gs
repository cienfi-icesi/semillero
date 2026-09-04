/**
 * Semillero CIENFI · Backend de avances
 * Guarda cada reporte en una hoja de cálculo y envía una notificación por correo.
 *
 * CONFIGURACIÓN (2 constantes)
 *  1. CORREO   → a dónde llegan las notificaciones.
 *  2. HOJA_ID  → el id de la hoja de cálculo donde se escriben los avances.
 *
 *     Cómo sacar el HOJA_ID: abre la hoja en Google Drive y copia el tramo
 *     que va entre /d/ y /edit en la barra de direcciones.
 *
 *     https://docs.google.com/spreadsheets/d/AQUI_VA_EL_ID/edit#gid=0
 *                                            ^^^^^^^^^^^^^
 *
 *     Si este script vive DENTRO de la hoja (Extensiones → Apps Script),
 *     puedes dejar HOJA_ID vacío y usará la hoja que lo contiene.
 *
 * DESPUÉS DE EDITAR
 *  - Ejecuta una vez la función `probar` desde el editor: eso autoriza los
 *    permisos de hoja de cálculo y de correo, y escribe una fila de prueba.
 *  - Implementar → Administrar implementaciones → editar → Implementar,
 *    para que la versión publicada tome los cambios. La URL no cambia.
 */

const CORREO  = 'efmartinez@icesi.edu.co';
const HOJA_ID = '1XjYRymtla-ea7F7M8EJ_Y6t_3CThLdjQ6hkH7atdCVQ';
const HOJA     = 'Avances';
const HOJA_ARCH = 'Fichas';
const CARPETA_ID = '';   // opcional: id de la carpeta de Drive para las fichas.
                         // Si se deja vacío, el script crea "Semillero CIENFI · fichas" en Mi unidad.

const COLS = ['Recibido','Equipo','Nombre del equipo','Línea','Integrantes','Paper','Autores','DOI',
              'Metodología','Lectura','Identificación','Infografía','% avance','Responsable',
              'Fecha reportada','Tareas pendientes','Observaciones'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.tipo === 'archivo') return guardarArchivo_(d);
    const sh = hoja_();
    sh.appendRow([
      new Date(), d.equipo, d.nombre, d.linea, d.integrantes, d.paper, d.paper_autores,
      d.paper_doi, d.paper_metodo, d.lectura, d.identificacion, d.infografia,
      d.porcentaje + '%', d.responsable, d.fecha, d.pendientes, d.observaciones
    ]);
    notificar_(d);
    return json_({ ok: true, fila: sh.getLastRow() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// Permite abrir la URL en el navegador para comprobar QUÉ VERSIÓN está publicada.
// Si al abrir la URL no ves "version":"v2", la implementación sigue sirviendo
// código viejo: Implementar → Administrar implementaciones → lápiz →
// Versión: Nueva versión → Implementar.
function doGet() {
  return json_({ ok: true, servicio: 'Semillero CIENFI · avances', version: 'v3' });
}

/** Ejecuta esto una vez desde el editor para autorizar permisos y probar. */
function probar() {
  const sh = hoja_();
  sh.appendRow([new Date(), 'TEST', 'Fila de prueba — se puede borrar']);
  MailApp.sendEmail(CORREO, 'Semillero · prueba de conexión',
                    'Si recibes este correo, el backend quedó bien configurado.');
  Logger.log('Hoja: ' + libro_().getUrl());
  Logger.log('Carpeta de fichas: ' + carpeta_().getUrl());
}

/** Recibe la ficha de lectura en base64, la guarda en Drive y la registra. */
function guardarArchivo_(d) {
  const bytes = Utilities.base64Decode(d.contenido);
  const nombre = d.equipo + ' · ' + d.archivo;
  const blob = Utilities.newBlob(bytes, d.mime || 'application/pdf', nombre);
  const archivo = carpeta_().createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = archivo.getUrl();

  hojaArch_().appendRow([new Date(), d.equipo, d.nombre_equipo, d.integrantes,
                         d.archivo, d.responsable, url]);

  MailApp.sendEmail(CORREO,
    'Semillero · ficha de lectura ' + d.equipo + ' — ' + d.archivo,
    'El equipo ' + d.equipo + ' (' + d.nombre_equipo + ') subió su ficha de lectura.\n\n' +
    'Integrantes: ' + d.integrantes + '\n' +
    'Responsable: ' + (d.responsable || '—') + '\n' +
    'Archivo: ' + d.archivo + '\n\n' + url);

  return json_({ ok: true, url: url });
}

function carpeta_() {
  if (CARPETA_ID) return DriveApp.getFolderById(CARPETA_ID);
  const nombre = 'Semillero CIENFI · fichas';
  const it = DriveApp.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : DriveApp.createFolder(nombre);
}

function hojaArch_() {
  const ss = libro_();
  let sh = ss.getSheetByName(HOJA_ARCH);
  if (!sh) {
    sh = ss.insertSheet(HOJA_ARCH);
    sh.appendRow(['Recibido','Equipo','Nombre del equipo','Integrantes','Archivo','Responsable','Enlace']);
    sh.getRange(1, 1, 1, 7).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function libro_() {
  const ss = HOJA_ID ? SpreadsheetApp.openById(HOJA_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No hay hoja de cálculo. Este script no está dentro de una hoja: ' +
                    'define HOJA_ID con el id de la hoja donde quieres guardar los avances.');
  }
  return ss;
}

function hoja_() {
  const ss = libro_();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.appendRow(COLS);
    sh.getRange(1, 1, 1, COLS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function notificar_(d) {
  const asunto = 'Semillero · avance ' + d.equipo + ' (' + d.porcentaje + '%) — ' +
                 (d.responsable || 'sin responsable');
  const cuerpo =
    'Equipo ' + d.equipo + ' — ' + d.nombre + '\n' +
    'Línea: ' + d.linea + '\n' +
    'Integrantes: ' + d.integrantes + '\n\n' +
    'Paper: ' + (d.paper || '—') + '\n' +
    (d.paper_doi ? 'DOI: https://doi.org/' + d.paper_doi + '\n' : '') +
    (d.paper_metodo ? 'Metodología: ' + d.paper_metodo + '\n' : '') + '\n' +
    'Lectura: ' + d.lectura + '\n' +
    'Identificación: ' + d.identificacion + '\n' +
    'Infografía: ' + d.infografia + '\n' +
    'Avance total: ' + d.porcentaje + '%\n\n' +
    'Responsable: ' + (d.responsable || '—') + '\n' +
    'Fecha reportada: ' + (d.fecha || '—') + '\n\n' +
    'Tareas pendientes:\n' + (d.pendientes || '—') + '\n\n' +
    'Observaciones:\n' + (d.observaciones || '—') + '\n\n' +
    '— Enviado desde el aplicativo del semillero\n' + libro_().getUrl();
  MailApp.sendEmail(CORREO, asunto, cuerpo);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
