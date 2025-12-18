/**
 * dividir.js
 * Requisitos:
 *   npm i xlsx inquirer@8 exceljs
 */

const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const readline = require('readline');
const fs = require('fs');
const path = require('path');
const inquirer = require('inquirer');

// ===============================
// CONFIGURACIÓN
// ===============================
const CHUNK_SIZE = 50;
const SHEET_INDEX = 0;
const OUTPUT_DIR = 'exportados';

const ROOT_DIR = __dirname;

const EXCLUDED_DIRS = new Set([
  'exportados',
  'node_modules',
  '.git'
]);

const DEBUG = true;

// Verde suave tipo Excel (EXPORTADO)
const MARK_FILL_OK = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFC6EFCE' }
};

// Amarillo suave (EN RANGO PERO SIN TELÉFONO / VACÍO)
const MARK_FILL_WARN = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFFEB9C' }
};

// Rojo suave (EN RANGO PERO TELÉFONO INVÁLIDO / NO CUMPLE)
const MARK_FILL_BAD = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFFC7CE' }
};

// Nueva columna de control
const CONTACT_COL_HEADER = 'C_CONTACTADO';

// ===============================
// PREPARAR CARPETA DE SALIDA
// ===============================
const OUT_DIR_ABS = path.join(ROOT_DIR, OUTPUT_DIR);

if (!fs.existsSync(OUT_DIR_ABS)) {
  fs.mkdirSync(OUT_DIR_ABS, { recursive: true });
  if (DEBUG) console.log(`[LOG] Carpeta creada: ${OUT_DIR_ABS}`);
} else {
  if (DEBUG) console.log(`[LOG] Carpeta de salida OK: ${OUT_DIR_ABS}`);
}

// ===============================
// UTILIDADES
// ===============================
function log(...args) { if (DEBUG) console.log('[LOG]', ...args); }
function warn(...args) { console.log('[WARN]', ...args); }
function err(...args) { console.error('[ERROR]', ...args); }

function extractPhoneNumbers(rawValue) {
  if (rawValue === null || rawValue === undefined) return [];

  const raw = String(rawValue).trim();
  if (!raw) return [];

  let normalized = raw
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\s+[-–—]\s+/g, '\n')
    .replace(/[,;|\/\\\t]+/g, '\n')
    .replace(/\s+(y|o)\s+/gi, '\n');

  const parts = normalized
    .split('\n')
    .map(s => s.trim())
    .filter(Boolean);

  const phones = [];
  for (const part of parts) {
    const digits = part.replace(/\D+/g, '');
    if (digits.length >= 7) phones.push(digits);
  }

  const seen = new Set();
  return phones.filter(p => {
    if (seen.has(p)) return false;
    seen.add(p);
    return true;
  });
}

function getDateStr() {
  const now = new Date();
  const day = String(now.getDate()).padStart(2, '0');
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const year = now.getFullYear();
  return `${day}-${month}-${year}`;
}

function normalizeHeader(h) {
  return String(h ?? '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}

function buildExportRows(records) {
  const out = [];

  for (const rec of records) {
    const baseName = String(rec.name ?? '').trim();
    const phones = extractPhoneNumbers(rec.phoneRaw);

    if (phones.length === 0) continue;

    for (let i = 0; i < phones.length; i++) {
      const suffix = i === 0 ? '' : ` (${i + 1})`;
      const nome = baseName ? `${baseName}${suffix}` : `Sin nombre${suffix}`;
      out.push([nome, phones[i], '']);
    }
  }

  return out;
}

function findNameColumn(headersOriginal) {
  const headersNorm = headersOriginal.map(h => normalizeHeader(h));

  const exactIdx = headersNorm.indexOf('nombre');
  if (exactIdx !== -1) return { idx: exactIdx, reason: 'match exacto: "NOMBRE"' };

  for (let i = 0; i < headersNorm.length; i++) {
    const hn = headersNorm[i];
    if (!hn) continue;
    if (hn.includes('rfc')) continue;
    if (hn.includes('nombre')) return { idx: i, reason: 'contiene "NOMBRE"' };
  }

  const strongSynonyms = [
    'razon social',
    'nombre o razon social',
    'cliente',
    'empresa',
    'contacto',
    'responsable',
    'titular',
  ];

  for (const syn of strongSynonyms) {
    const idx = headersNorm.indexOf(syn);
    if (idx !== -1) return { idx, reason: `sinónimo fuerte: "${syn}"` };
  }

  for (let i = 0; i < headersNorm.length; i++) {
    const hn = headersNorm[i];
    if (!hn) continue;
    if (hn.includes('rfc')) continue;

    const partialHits = [
      'razon social',
      'cliente',
      'empresa',
      'contacto',
      'responsable',
      'titular',
    ];

    for (const p of partialHits) {
      if (hn.includes(p)) return { idx: i, reason: `coincidencia parcial: "${p}"` };
    }
  }

  return null;
}

function findPhoneColumn(headersOriginal) {
  const headersNorm = headersOriginal.map(h => normalizeHeader(h));
  const blacklist = ['correo', 'email', 'e-mail', 'mail'];

  const priorityExact = ['telefono', 'teléfono', 'celular', 'whatsapp', 'movil', 'móvil', 'tel'];
  for (const p of priorityExact) {
    const idx = headersNorm.indexOf(normalizeHeader(p));
    if (idx !== -1) {
      const hn = headersNorm[idx];
      if (blacklist.some(b => hn.includes(b))) continue;
      return { idx, reason: `match exacto: "${p}"` };
    }
  }

  const priorityContains = ['telefono', 'celular', 'whatsapp', 'movil', 'tel', 'contacto'];
  for (let i = 0; i < headersNorm.length; i++) {
    const hn = headersNorm[i];
    if (!hn) continue;
    if (blacklist.some(b => hn.includes(b))) continue;

    for (const p of priorityContains) {
      if (hn.includes(p)) return { idx: i, reason: `contiene: "${p}"` };
    }
  }

  return null;
}

function detectColumns(headers) {
  const nameInfo = findNameColumn(headers);
  const phoneInfo = findPhoneColumn(headers);

  if (!nameInfo || !phoneInfo) return null;

  return {
    profile: 'auto',
    idxName: nameInfo.idx,
    idxPhone: phoneInfo.idx,
    nameReason: nameInfo.reason,
    phoneReason: phoneInfo.reason,
  };
}

function excelRowToIndex(excelRow) { return excelRow - 2; }
function indexToExcelRow(index) { return index + 2; }

function askYesNo(question) {
  return new Promise(resolve => {
    const r2 = readline.createInterface({ input: process.stdin, output: process.stdout });
    r2.question(question, ans => {
      r2.close();
      const v = String(ans || '').trim().toUpperCase();
      resolve(v === 'S' || v === 'SI' || v === 'SÍ' || v === 'Y' || v === 'YES');
    });
  });
}

// ===============================
// EXPLORADOR
// ===============================
function isExcelFile(name) {
  const lower = name.toLowerCase();
  return (lower.endsWith('.xlsx') || lower.endsWith('.xls') || lower.endsWith('.xlsm')) && !lower.startsWith('~$');
}

function safeReadDir(dir) {
  try {
    return fs.readdirSync(dir, { withFileTypes: true });
  } catch (e) {
    warn(`No pude leer directorio: ${dir}`);
    warn(String(e?.message || e));
    return [];
  }
}

function listBrowserChoices(currentDir) {
  const entries = safeReadDir(currentDir);

  const dirs = [];
  const excels = [];

  for (const e of entries) {
    if (e.isDirectory()) {
      if (EXCLUDED_DIRS.has(e.name)) continue;
      dirs.push(e.name);
    } else if (e.isFile()) {
      if (!isExcelFile(e.name)) continue;
      excels.push(e.name);
    }
  }

  dirs.sort((a, b) => a.localeCompare(b));
  excels.sort((a, b) => a.localeCompare(b));

  const choices = [];

  if (path.resolve(currentDir) !== path.resolve(ROOT_DIR)) {
    choices.push({ name: '⬅️  [..] Subir un nivel', value: { type: 'up' } });
  }

  for (const d of dirs) {
    choices.push({ name: `📁  ${d}`, value: { type: 'dir', name: d } });
  }

  if (excels.length > 0) {
    for (const f of excels) {
      const full = path.join(currentDir, f);
      let meta = '';
      try {
        const st = fs.statSync(full);
        const kb = Math.max(1, Math.round(st.size / 1024));
        meta = `  —  ${kb} KB  —  ${st.mtime.toLocaleString()}`;
      } catch {}
      choices.push({ name: `📄  ${f}${meta}`, value: { type: 'file', name: f } });
    }
  } else {
    choices.push({ name: '⚠️  (No hay Excel aquí; entra a otra carpeta)', value: { type: 'noop' }, disabled: true });
  }

  return { choices, dirsCount: dirs.length, excelsCount: excels.length };
}

async function pickExcelFile() {
  let currentDir = ROOT_DIR;

  log('ROOT_DIR =', ROOT_DIR);
  try {
    const preview = fs.readdirSync(ROOT_DIR).slice(0, 30);
    log('Contenido ROOT_DIR (preview 30):', preview);
  } catch {}

  while (true) {
    const rel = path.relative(ROOT_DIR, currentDir) || '.';
    const { choices, dirsCount, excelsCount } = listBrowserChoices(currentDir);

    log(`Explorando carpeta: ${currentDir} (rel: ${rel}) | dirs=${dirsCount} excels=${excelsCount}`);

    try {
      const ans = await inquirer.prompt([
        {
          type: 'list',
          name: 'pick',
          message: `📄 Selecciona el Excel (↑ ↓ Enter) | Carpeta: ${rel}`,
          choices,
          pageSize: 14
        }
      ]);

      const pick = ans.pick;

      if (pick.type === 'up') {
        currentDir = path.dirname(currentDir);
        continue;
      }

      if (pick.type === 'dir') {
        currentDir = path.join(currentDir, pick.name);
        continue;
      }

      if (pick.type === 'file') {
        return path.join(currentDir, pick.name);
      }
    } catch (e) {
      err('Se canceló la selección o hubo un error del prompt.');
      err(String(e?.message || e));
      process.exit(1);
    }
  }
}

// ===============================
// MARCADO + C_CONTACTADO (EN UNA SOLA APERTURA DE EXCEL)
// ===============================
async function applyMarksAndContactado({
  inputFile,
  sheetName,
  colNumbersToMark,
  okRows,
  warnRows,
  badRows
}) {
  log('Aplicando marcado y C_CONTACTADO...');
  log('Archivo origen =', inputFile);
  log('Hoja =', sheetName);

  const cols = [...new Set(colNumbersToMark)]
    .map(n => parseInt(n, 10))
    .filter(n => Number.isFinite(n) && n >= 1);

  if (cols.length === 0) {
    warn('No hay columnas válidas para marcar (colNumbersToMark vacío).');
    return;
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(inputFile);

  const ws = wb.getWorksheet(sheetName) || wb.worksheets[SHEET_INDEX];
  if (!ws) throw new Error('Worksheet not found');

  // 1) Asegurar columna "C_CONTACTADO" (si no existe)
  const headerRow = ws.getRow(1);
  let contactColIndex = null;

  for (let c = 1; c <= ws.columnCount; c++) {
    const v = headerRow.getCell(c).value;
    const txt = String(v ?? '').trim();
    if (txt.toUpperCase() === CONTACT_COL_HEADER) {
      contactColIndex = c;
      break;
    }
  }

  if (!contactColIndex) {
    contactColIndex = ws.columnCount + 1;
    headerRow.getCell(contactColIndex).value = CONTACT_COL_HEADER;
    headerRow.commit();
    log(`Columna agregada: ${CONTACT_COL_HEADER} en columna #${contactColIndex}`);
  } else {
    log(`Columna existente: ${CONTACT_COL_HEADER} en columna #${contactColIndex}`);
  }

  // 2) Helper: aplicar fill (solo si no tenía fill) + setear SI/NO en C_CONTACTADO
  // Nota: por "historial", NO sobrescribimos C_CONTACTADO si ya tiene valor.
  function processRowSet(rowNumbers, fill, contactValue) {
    for (const r of rowNumbers) {
      if (r < 2) continue;

      const row = ws.getRow(r);

      // Marcar SOLO columnas afectadas (ej. Nombre y Teléfono)
      for (const c of cols) {
        const cell = row.getCell(c);
        if (cell.fill) continue; // respeta marcas previas
        cell.fill = fill;
      }

      // Escribir SI/NO en C_CONTACTADO (solo si está vacío)
      const contactCell = row.getCell(contactColIndex);
      const current = String(contactCell.value ?? '').trim();
      if (!current) {
        contactCell.value = contactValue; // "SI" o "NO"
      }

      row.commit();
    }
  }

  // Orden recomendado:
  // OK primero, luego WARN y BAD (como respetamos fills previos, no se pisan)
  processRowSet(okRows, MARK_FILL_OK, 'SI');
  processRowSet(warnRows, MARK_FILL_WARN, 'NO');
  processRowSet(badRows, MARK_FILL_BAD, 'NO');

  await wb.xlsx.writeFile(inputFile);
  log('Marcado + C_CONTACTADO: COMPLETADO (archivo sobrescrito).');
}

// ===============================
// EXPORTACIÓN
// ===============================
async function exportChunks(expandedRows, startIndex, endExclusiveIndex, inputBaseName) {
  const dateStr = getDateStr();
  const excelStart = indexToExcelRow(startIndex);
  const excelEnd = indexToExcelRow(endExclusiveIndex - 1);

  const totalFiles = Math.ceil(expandedRows.length / CHUNK_SIZE);
  log('Export totalFiles =', totalFiles, '| CHUNK_SIZE =', CHUNK_SIZE);

  const expectedPaths = [];
  for (let k = 1; k <= totalFiles; k++) {
    const fileName = `${inputBaseName}_${dateStr}_(${excelStart}-${excelEnd})_${String(k).padStart(2, '0')}.xlsx`;
    expectedPaths.push(path.join(OUT_DIR_ABS, fileName));
  }

  const anyCollision = expectedPaths.some(p => fs.existsSync(p));
  log('Colisión detectada =', anyCollision);

  let overwriteExisting = false;
  if (anyCollision) {
    overwriteExisting = await askYesNo('⚠️ Ya existen archivos con ese origen/fecha/rango. ¿Deseas sobrescribirlos? (S/N): ');
    log('overwriteExisting =', overwriteExisting);
  }

  let fileCounter = 1;

  for (let i = 0; i < expandedRows.length; i += CHUNK_SIZE) {
    const chunk = expandedRows.slice(i, i + CHUNK_SIZE);

    const sheetData = [
      ['nome', 'numero', 'e-mail'],
      ...chunk
    ];

    const newWb = XLSX.utils.book_new();
    const newSheet = XLSX.utils.aoa_to_sheet(sheetData);
    XLSX.utils.book_append_sheet(newWb, newSheet, 'Hoja1');

    let outputPath;
    while (true) {
      const fileName = `${inputBaseName}_${dateStr}_(${excelStart}-${excelEnd})_${String(fileCounter).padStart(2, '0')}.xlsx`;
      outputPath = path.join(OUT_DIR_ABS, fileName);

      if (overwriteExisting) break;
      if (!fs.existsSync(outputPath)) break;

      fileCounter++;
    }

    XLSX.writeFile(newWb, outputPath);
    console.log(`✅ Generado: ${outputPath} (${chunk.length} registros)`);

    fileCounter++;
  }

  console.log('🎉 Proceso completado con éxito.');
}

// ===============================
// MAIN
// ===============================
(async () => {
  const INPUT_FILE = await pickExcelFile();
  const inputBaseName = path.parse(INPUT_FILE).name;

  log('INPUT_FILE =', INPUT_FILE);
  log('inputBaseName =', inputBaseName);

  const wb = XLSX.readFile(INPUT_FILE);
  const sheetName = wb.SheetNames[SHEET_INDEX];
  if (!sheetName) throw new Error('No sheets found in Excel.');

  const sheet = wb.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  if (!data || data.length === 0) {
    err('El archivo está vacío o no se pudo leer.');
    process.exit(1);
  }

  const headers = data[0].map(h => String(h ?? '').trim());
  log('Encabezados detectados:', headers);

  const detected = detectColumns(headers);
  if (!detected) {
    err('No se pudieron detectar columnas requeridas por encabezados.');
    console.log('Encabezados encontrados:', headers);
    process.exit(1);
  }

  const idxName = detected.idxName;
  const idxPhone = detected.idxPhone;

  // Columnas reales (Excel 1-based) para marcar SOLO lo “afectado”
  const colsToMark = [idxName + 1, idxPhone + 1];

  console.log(`✅ Archivo seleccionado: ${INPUT_FILE} (perfil: ${detected.profile})`);
  console.log(`🧩 Columna NOMBRE detectada: "${headers[idxName]}" → ${detected.nameReason}`);
  console.log(`📞 Columna TEL/CEL detectada: "${headers[idxPhone]}" → ${detected.phoneReason}`);
  console.log(`🟩🟨🟥 Se marcarán SOLO estas columnas (1-based): ${colsToMark.join(', ')}`);
  console.log(`🧾 Se creará/actualizará columna: ${CONTACT_COL_HEADER} (SI/NO)`);

  // Guardar excelRowNumber real por registro
  const baseRecords = data
    .slice(1)
    .map((row, i) => {
      const excelRowNumber = i + 2; // fila real en Excel
      const name = row[idxName] ?? '';
      const phoneRaw = row[idxPhone] ?? '';
      return { name, phoneRaw, excelRowNumber };
    })
    .filter(r => String(r.name).trim() !== '' || String(r.phoneRaw).trim() !== '');

  if (baseRecords.length === 0) {
    err('No hay datos válidos para procesar.');
    process.exit(1);
  }

  const totalFilasExcelDatos = baseRecords.length + 1;
  console.log(`ℹ️ Total de filas con datos (según Excel): ${baseRecords.length}.`);
  console.log('   Recuerda: la fila 2 de Excel es la primera fila con datos.');

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

  rl.question('👉 ¿Desde qué número de fila de EXCEL quieres empezar? (ej. 2, 1400): ', answerStart => {
    let excelRowStart = parseInt(String(answerStart).trim(), 10);

    if (Number.isNaN(excelRowStart) || excelRowStart < 2) {
      warn('Inicio no válido. Se usará la primera fila con datos (fila 2).');
      excelRowStart = 2;
    }

    const startIndex = excelRowToIndex(excelRowStart);
    log('excelRowStart =', excelRowStart, '| startIndex =', startIndex);

    if (startIndex >= baseRecords.length) {
      warn(`La fila inicial (${excelRowStart}) está fuera del rango de datos (máximo ${totalFilasExcelDatos}). No hay nada que procesar.`);
      rl.close();
      process.exit(0);
    }

    rl.question(`👉 ¿Hasta qué número de fila de EXCEL quieres procesar? (Enter = hasta la última, máximo ${totalFilasExcelDatos}): `, answerEnd => {
      const trimmed = String(answerEnd).trim();
      let endExclusiveIndex;

      if (trimmed === '') {
        endExclusiveIndex = baseRecords.length;
        console.log(`✅ Se procesará desde la fila ${excelRowStart} hasta la última fila con datos (${totalFilasExcelDatos}).`);
      } else {
        let excelRowEnd = parseInt(trimmed, 10);

        if (Number.isNaN(excelRowEnd)) {
          warn('Final no válido. Se usará hasta la última fila con datos.');
          endExclusiveIndex = baseRecords.length;
        } else {
          if (excelRowEnd < excelRowStart) {
            warn('La fila final es menor que la inicial. No se procesará nada.');
            rl.close();
            process.exit(0);
          }

          if (excelRowEnd > totalFilasExcelDatos) {
            warn(`La fila final (${excelRowEnd}) excede el máximo (${totalFilasExcelDatos}). Se ajustará al máximo.`);
            excelRowEnd = totalFilasExcelDatos;
          }

          endExclusiveIndex = excelRowEnd - 1;

          if (endExclusiveIndex <= startIndex) {
            warn('El rango calculado está vacío. No hay nada que procesar.');
            rl.close();
            process.exit(0);
          }

          if (endExclusiveIndex > baseRecords.length) {
            endExclusiveIndex = baseRecords.length;
          }

          console.log(`✅ Se procesará desde la fila ${excelRowStart} hasta la fila ${excelRowEnd} de Excel.`);
        }
      }

      rl.close();

      const rangedRecords = baseRecords.slice(startIndex, endExclusiveIndex);

      // ===============================
      // CLASIFICACIÓN DE FILAS EN RANGO
      // ===============================
      // OK (verde): tiene ≥1 teléfono válido (>=7 dígitos)
      // WARN (amarillo): teléfono vacío / nulo
      // BAD (rojo): hay contenido, pero NO produce teléfonos válidos
      const okRows = [];
      const warnRows = [];
      const badRows = [];

      for (const rec of rangedRecords) {
        const raw = String(rec.phoneRaw ?? '').trim();
        const phones = extractPhoneNumbers(rec.phoneRaw);

        if (phones.length > 0) {
          okRows.push(rec.excelRowNumber);
        } else if (!raw) {
          warnRows.push(rec.excelRowNumber);
        } else {
          badRows.push(rec.excelRowNumber);
        }
      }

      // Exportación (solo depende de los registros)
      const expandedRows = buildExportRows(rangedRecords);

      log('rangedRecords.length =', rangedRecords.length);
      log('expandedRows.length (post limpieza/expansión) =', expandedRows.length);
      log('okRows =', okRows.length, '| warnRows =', warnRows.length, '| badRows =', badRows.length);

      if (expandedRows.length === 0) {
        warn('No hubo teléfonos válidos para exportar en ese rango.');
        // Aun así, marcamos WARN/BAD y ponemos NO en C_CONTACTADO, porque hubo rango procesado
        applyMarksAndContactado({
          inputFile: INPUT_FILE,
          sheetName,
          colNumbersToMark: colsToMark,
          okRows: [],
          warnRows,
          badRows
        }).then(() => {
          console.log('✅ Se marcó el rango (sin exportación) y se registró C_CONTACTADO=NO donde aplicaba.');
          process.exit(0);
        }).catch(e => {
          err('Error al marcar/registrar:');
          err(String(e?.message || e));
          process.exit(1);
        });
        return;
      }

      exportChunks(expandedRows, startIndex, endExclusiveIndex, inputBaseName)
        .then(async () => {
          console.log('🟩🟨🟥 Aplicando marcado + C_CONTACTADO...');
          console.log(`   🟩 OK(exportado): ${okRows.length} filas`);
          console.log(`   🟨 WARN(vacío):   ${warnRows.length} filas`);
          console.log(`   🟥 BAD(inválido): ${badRows.length} filas`);
          console.log(`   🧾 ${CONTACT_COL_HEADER}: SI para verde / NO para amarillo y rojo`);

          // Nota: si el archivo está abierto en Excel, esto puede fallar (bloqueo).
          await applyMarksAndContactado({
            inputFile: INPUT_FILE,
            sheetName,
            colNumbersToMark: colsToMark,
            okRows,
            warnRows,
            badRows
          });

          console.log('✅ Marcado y C_CONTACTADO completados (Excel origen sobrescrito, marcas previas respetadas).');
        })
        .catch(e => {
          err('Error al exportar o al marcar/registrar:');
          err(String(e?.message || e));
          process.exit(1);
        });
    });
  });
})();
