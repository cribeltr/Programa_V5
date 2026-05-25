/**
 * PMP 2026 - Backend Google Apps Script
 * Hospital Hernán Henríquez Aravena
 *
 * INSTRUCCIONES DE DESPLIEGUE:
 * 1. Crear un Google Sheet nuevo y vacío.
 * 2. Extensiones → Apps Script. Borrar el contenido por defecto.
 * 3. Pegar este archivo completo. Guardar.
 * 4. Ejecutar la función setup() una vez (botón Run). Autorizar los permisos.
 * 5. Desplegar → Nueva implementación → Tipo: Aplicación web.
 *    - Ejecutar como: Yo.
 *    - Quién tiene acceso: Cualquier usuario (o "Cualquiera con el enlace").
 * 6. Copiar la URL del Web App.
 * 7. En el index.html → Configuración: pegar el ID del Sheet y la URL del Web App.
 */

const SCHEMA_VERSION = '1.0.0';

const SHEETS = {
  equipos: ['id','n_carpeta','n_inventario','equipo','servicio','unidad','ubicacion','procedencia','marca','modelo','serie','anio_instalacion','vida_util_residual','clasificacion','enu_baja','frecuencia_mp','es_slot_disponible','estado_actual','responsable_actual','updated_at'],
  programacion: ['equipo_id','mes','codigo','updated_at'],
  registro: ['equipo_id','mes','programacion','resultado','ejecutor','updated_at'],
  eventos: ['id','equipo_id','tipo','fecha_evento','payload_json','created_at'],
  pendientes: ['id','equipo_id','origen','descripcion','ejecutor_asignado','fecha_creacion','fecha_compromiso','estado','ultima_actualizacion'],
  pendientes_tareas: ['id','pendiente_id','descripcion','estado','created_at'],
  pendientes_actualizaciones: ['id','pendiente_id','fecha','texto','created_at'],
  asignaciones: ['equipo_id','mes','responsable','updated_at'],
  meta: ['clave','valor','updated_at']
};

function setup() {
  ensureSheets_();
  setMeta_('schema_version', SCHEMA_VERSION);
  setMeta_('setup_at', new Date().toISOString());
  SpreadsheetApp.getUi().alert('Setup completado. Pestañas creadas/verificadas: ' + Object.keys(SHEETS).join(', '));
}

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      const headers = SHEETS[name];
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.autoResizeColumns(1, headers.length);
    } else {
      // Verificar que tiene encabezados; si está vacío, agregarlos
      if (sh.getLastRow() === 0) {
        const headers = SHEETS[name];
        sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
        sh.setFrozenRows(1);
      }
    }
  });
  // Borrar Sheet1 por defecto si está vacío y no es uno de los nuestros
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja 1') || ss.getSheetByName('Hoja1');
  if (def && !SHEETS[def.getName()] && ss.getSheets().length > 1 && def.getLastRow() <= 1) {
    try { ss.deleteSheet(def); } catch (e) {}
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    ensureSheets_();
    const action = (e && e.parameter && e.parameter.action) || 'ping';
    switch (action) {
      case 'ping': return json_({ok:true, pong:true, schema:SCHEMA_VERSION, time:new Date().toISOString()});
      case 'getAll': return json_({ok:true, data:{
        equipos: readSheet_('equipos'),
        programacion: readSheet_('programacion'),
        registro: readSheet_('registro'),
        eventos: readSheet_('eventos'),
        pendientes: readSheet_('pendientes'),
        pendientes_tareas: readSheet_('pendientes_tareas'),
        pendientes_actualizaciones: readSheet_('pendientes_actualizaciones'),
        asignaciones: readSheet_('asignaciones'),
        meta: readSheet_('meta')
      }});
      case 'getEquipos': return json_({ok:true, data:readSheet_('equipos')});
      case 'getProgramacion': return json_({ok:true, data:readSheet_('programacion')});
      case 'getRegistro': return json_({ok:true, data:readSheet_('registro')});
      case 'getEventos': return json_({ok:true, data:readSheet_('eventos')});
      case 'getPendientes': return json_({ok:true, data:{
        pendientes: readSheet_('pendientes'),
        tareas: readSheet_('pendientes_tareas'),
        actualizaciones: readSheet_('pendientes_actualizaciones')
      }});
      case 'getAsignaciones': return json_({ok:true, data:readSheet_('asignaciones')});
      case 'getMeta': return json_({ok:true, data:readSheet_('meta')});
      default: return json_({ok:false, error:'Acción desconocida: '+action});
    }
  } catch (err) {
    return json_({ok:false, error:String(err && err.message || err)});
  }
}

function doPost(e) {
  try {
    ensureSheets_();
    let body = {};
    if (e && e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch (er) { body = {}; }
    }
    const action = body.action || (e && e.parameter && e.parameter.action) || '';
    const payload = body.payload || {};
    switch (action) {
      case 'saveEquipos': return json_({ok:true, count:upsertEquipos_(payload.rows || [])});
      case 'saveProgramacion': return json_({ok:true, count:upsertByKey_('programacion', payload.rows || [], ['equipo_id','mes'])});
      case 'saveRegistro': return json_({ok:true, count:upsertByKey_('registro', payload.rows || [], ['equipo_id','mes'])});
      case 'saveAsignaciones': return json_({ok:true, count:upsertByKey_('asignaciones', payload.rows || [], ['equipo_id','mes'])});
      case 'addEvento': {
        const row = payload.row || {};
        row.id = row.id || Utilities.getUuid();
        row.created_at = new Date().toISOString();
        appendRow_('eventos', row);
        // Si el evento es mp y trae resultado, actualizar registro
        if (row.tipo === 'mp' && row.equipo_id) {
          try {
            const pl = typeof row.payload_json === 'string' ? JSON.parse(row.payload_json) : (row.payload_json || {});
            if (pl.mes && pl.resultado) {
              upsertByKey_('registro', [{
                equipo_id: row.equipo_id,
                mes: pl.mes,
                programacion: pl.programacion_original || 'X',
                resultado: pl.resultado,
                ejecutor: pl.ejecutor || '',
                updated_at: new Date().toISOString()
              }], ['equipo_id','mes']);
            }
            if (pl.estado_equipo) {
              updateEquipoEstado_(row.equipo_id, pl.estado_equipo === 'operativo' ? 'Operativo' : 'No operativo');
            }
          } catch(er) {}
        }
        if (row.tipo === 'envio_st' && row.equipo_id) updateEquipoEstado_(row.equipo_id, 'En servicio técnico');
        if (row.tipo === 'recepcion' && row.equipo_id) {
          try {
            const pl = typeof row.payload_json === 'string' ? JSON.parse(row.payload_json) : (row.payload_json || {});
            updateEquipoEstado_(row.equipo_id, pl.estado === 'operativo' ? 'Operativo' : 'No operativo');
          } catch(er) {}
        }
        if (row.tipo === 'reparacion' && row.equipo_id) {
          try {
            const pl = typeof row.payload_json === 'string' ? JSON.parse(row.payload_json) : (row.payload_json || {});
            updateEquipoEstado_(row.equipo_id, pl.estado === 'operativo' ? 'Operativo' : 'No operativo');
          } catch(er) {}
        }
        if (row.tipo === 'solicitud' && row.equipo_id) updateEquipoEstado_(row.equipo_id, 'No operativo');
        return json_({ok:true, id:row.id});
      }
      case 'addPendiente': {
        const row = payload.row || {};
        row.id = row.id || Utilities.getUuid();
        row.fecha_creacion = row.fecha_creacion || new Date().toISOString();
        row.ultima_actualizacion = new Date().toISOString();
        row.estado = row.estado || 'abierto';
        row.origen = row.origen || 'manual';
        appendRow_('pendientes', row);
        return json_({ok:true, id:row.id});
      }
      case 'updatePendiente': {
        const row = payload.row || {};
        row.ultima_actualizacion = new Date().toISOString();
        upsertByKey_('pendientes', [row], ['id']);
        return json_({ok:true});
      }
      case 'bulkUpdatePendientes': {
        const ids = payload.ids || [];
        const estado = payload.estado || '';
        const ejecutor = payload.ejecutor;
        if (!ids.length || !estado) return json_({ok:false, error:'Faltan ids o estado'});
        const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pendientes');
        const headers = SHEETS.pendientes;
        const last = sh.getLastRow();
        if (last < 2) return json_({ok:true, count:0});
        const range = sh.getRange(2, 1, last - 1, headers.length);
        const values = range.getValues();
        const idCol = headers.indexOf('id');
        const estCol = headers.indexOf('estado');
        const upCol = headers.indexOf('ultima_actualizacion');
        const ejCol = headers.indexOf('ejecutor_asignado');
        const idSet = new Set(ids.map(String));
        let count = 0;
        const now = new Date().toISOString();
        values.forEach(r => {
          if (idSet.has(String(r[idCol]))) {
            r[estCol] = estado;
            r[upCol] = now;
            if (ejecutor !== undefined && ejecutor !== null && ejecutor !== '') r[ejCol] = ejecutor;
            count++;
          }
        });
        range.setValues(values);
        return json_({ok:true, count});
      }
      case 'addTareaPendiente': {
        const row = payload.row || {};
        row.id = row.id || Utilities.getUuid();
        row.created_at = new Date().toISOString();
        row.estado = row.estado || 'abierto';
        appendRow_('pendientes_tareas', row);
        touchPendiente_(row.pendiente_id);
        return json_({ok:true, id:row.id});
      }
      case 'addActualizacionPendiente': {
        const row = payload.row || {};
        row.id = row.id || Utilities.getUuid();
        row.created_at = new Date().toISOString();
        appendRow_('pendientes_actualizaciones', row);
        touchPendiente_(row.pendiente_id);
        return json_({ok:true, id:row.id});
      }
      case 'applyDiff': {
        const diffs = payload.diffs || [];
        let creados = 0;
        diffs.forEach(d => {
          const row = {
            id: Utilities.getUuid(),
            equipo_id: d.equipo_id || '',
            origen: 'diff_carga',
            descripcion: d.descripcion || '',
            ejecutor_asignado: '',
            fecha_creacion: new Date().toISOString(),
            fecha_compromiso: '',
            estado: 'abierto',
            ultima_actualizacion: new Date().toISOString()
          };
          appendRow_('pendientes', row);
          creados++;
        });
        return json_({ok:true, count:creados});
      }
      case 'setMeta': {
        setMeta_(payload.clave, payload.valor);
        return json_({ok:true});
      }
      default: return json_({ok:false, error:'Acción desconocida: '+action});
    }
  } catch (err) {
    return json_({ok:false, error:String(err && err.message || err), stack:String(err && err.stack || '')});
  }
}

function readSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const headers = SHEETS[name];
  const values = sh.getRange(2, 1, last - 1, headers.length).getValues();
  return values.map(r => {
    const o = {};
    headers.forEach((h, i) => {
      let v = r[i];
      if (v instanceof Date) v = v.toISOString();
      o[h] = v === '' ? '' : v;
    });
    return o;
  }).filter(o => Object.values(o).some(v => v !== '' && v !== null && v !== undefined));
}

function appendRow_(name, obj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  const headers = SHEETS[name];
  const row = headers.map(h => obj[h] === undefined || obj[h] === null ? '' : obj[h]);
  sh.appendRow(row);
}

function upsertByKey_(name, rows, keyCols) {
  if (!rows.length) return 0;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  const headers = SHEETS[name];
  const last = sh.getLastRow();
  const existing = last >= 2 ? sh.getRange(2, 1, last - 1, headers.length).getValues() : [];
  const idx = new Map();
  existing.forEach((r, i) => {
    const k = keyCols.map(kc => String(r[headers.indexOf(kc)])).join('||');
    idx.set(k, i + 2);
  });
  const toAppend = [];
  const updates = [];
  rows.forEach(o => {
    const k = keyCols.map(kc => String(o[kc] === undefined ? '' : o[kc])).join('||');
    o.updated_at = o.updated_at || new Date().toISOString();
    const rowArr = headers.map(h => o[h] === undefined || o[h] === null ? '' : o[h]);
    if (idx.has(k)) {
      updates.push({rowNum: idx.get(k), values: rowArr});
    } else {
      toAppend.push(rowArr);
    }
  });
  updates.forEach(u => sh.getRange(u.rowNum, 1, 1, headers.length).setValues([u.values]));
  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, headers.length).setValues(toAppend);
  }
  return rows.length;
}

function upsertEquipos_(rows) {
  if (!rows.length) return 0;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('equipos');
  const headers = SHEETS.equipos;
  const last = sh.getLastRow();
  const existing = last >= 2 ? sh.getRange(2, 1, last - 1, headers.length).getValues() : [];
  const idx = new Map();
  existing.forEach((r, i) => {
    idx.set(String(r[0]), {rowNum: i + 2, row: r});
  });
  const toAppend = [];
  rows.forEach(o => {
    o.updated_at = new Date().toISOString();
    if (!o.id) o.id = Utilities.getUuid();
    const rowArr = headers.map(h => o[h] === undefined || o[h] === null ? '' : o[h]);
    if (idx.has(String(o.id))) {
      // Conservar estado_actual y responsable_actual si no vienen en el payload
      const ex = idx.get(String(o.id)).row;
      const i_estado = headers.indexOf('estado_actual');
      const i_resp = headers.indexOf('responsable_actual');
      if (!o.estado_actual && ex[i_estado]) rowArr[i_estado] = ex[i_estado];
      if (!o.responsable_actual && ex[i_resp]) rowArr[i_resp] = ex[i_resp];
      sh.getRange(idx.get(String(o.id)).rowNum, 1, 1, headers.length).setValues([rowArr]);
    } else {
      toAppend.push(rowArr);
    }
  });
  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, headers.length).setValues(toAppend);
  }
  return rows.length;
}

function updateEquipoEstado_(equipo_id, estado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('equipos');
  const headers = SHEETS.equipos;
  const last = sh.getLastRow();
  if (last < 2) return;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(equipo_id)) {
      const col = headers.indexOf('estado_actual') + 1;
      sh.getRange(i + 2, col).setValue(estado);
      sh.getRange(i + 2, headers.indexOf('updated_at') + 1).setValue(new Date().toISOString());
      return;
    }
  }
}

function touchPendiente_(pendiente_id) {
  if (!pendiente_id) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('pendientes');
  const headers = SHEETS.pendientes;
  const last = sh.getLastRow();
  if (last < 2) return;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(pendiente_id)) {
      sh.getRange(i + 2, headers.indexOf('ultima_actualizacion') + 1).setValue(new Date().toISOString());
      return;
    }
  }
}

function setMeta_(clave, valor) {
  if (!clave) return;
  upsertByKey_('meta', [{clave: clave, valor: typeof valor === 'object' ? JSON.stringify(valor) : String(valor), updated_at: new Date().toISOString()}], ['clave']);
}
