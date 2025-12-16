const XLSX = require('xlsx');
const readline = require('readline');
const fs = require('fs');
const path = require('path');

// ===============================
// CONFIGURACIÓN
// ===============================
const INPUT_FILE = 'clientes.xlsx';   // Archivo original
const CHUNK_SIZE = 50;                // Cantidad de registros por archivo
const SHEET_INDEX = 0;                // Primera hoja del Excel
const OUTPUT_DIR = 'exportados';      // Carpeta donde se guardarán los archivos exportados

// ===============================
// PREPARAR CARPETA DE SALIDA
// ===============================
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

// ===============================
// LECTURA DEL ARCHIVO
// ===============================
const wb = XLSX.readFile(INPUT_FILE);
const sheetName = wb.SheetNames[SHEET_INDEX];
const sheet = wb.Sheets[sheetName];

const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

if (!data || data.length === 0) {
  console.error('❌ El archivo está vacío o no se pudo leer.');
  process.exit(1);
}

// ===============================
// IDENTIFICAR COLUMNAS DEL ORIGINAL
// ===============================
const headers = data[0];

const COL_NOMBRE = 'Nombre';
const COL_TELEFONO = 'Telefono';

const idxNombre = headers.indexOf(COL_NOMBRE);
const idxTelefono = headers.indexOf(COL_TELEFONO);

if (idxNombre === -1 || idxTelefono === -1) {
  console.error('❌ No se encontraron las columnas "Nombre" y/o "Telefono".');
  console.log('Encabezados encontrados:', headers);
  process.exit(1);
}

// ===============================
// FILTRAR SOLO DATOS RELEVANTES
// ===============================
// rows[0] corresponde a la fila 2 de Excel (la primera con datos)
const rows = data
  .slice(1) // Quitar encabezados
  .map(row => {
    const nombre = row[idxNombre] ?? '';
    const telefono = row[idxTelefono] ?? '';
    return [nombre, telefono];
  })
  .filter(([nombre, telefono]) => {
    return String(nombre).trim() !== '' || String(telefono).trim() !== '';
  });

if (rows.length === 0) {
  console.error('❌ No hay datos válidos para procesar.');
  process.exit(1);
}

const totalFilasExcel = rows.length + 1; // porque fila 2 es rows[0]

console.log(`ℹ️ Total de filas con datos: ${rows.length}.`);
console.log('   Recuerda: la fila 2 de Excel es la primera fila con datos.');

// ===============================
// PREGUNTAR RANGO AL USUARIO
// ===============================
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

rl.question('👉 ¿Desde qué número de fila de EXCEL quieres empezar? (ej. 2, 1400): ', answerStart => {
  let excelRowStart = parseInt(answerStart.trim(), 10);
  let startIndex = 0;

  if (Number.isNaN(excelRowStart)) {
    console.log('⚠️ Entrada no válida. Se usará la primera fila con datos (fila 2 en Excel).');
    excelRowStart = 2;
  }

  if (excelRowStart < 2) {
    console.log('⚠️ La fila mínima con datos es la 2. Se usará la primera fila con datos.');
    excelRowStart = 2;
  }

  // Convertir número de fila de Excel a índice interno
  startIndex = excelRowStart - 2;

  if (startIndex >= rows.length) {
    console.log(`⚠️ La fila inicial (${excelRowStart}) está fuera del rango de datos (máximo ${totalFilasExcel}). No hay nada que procesar.`);
    rl.close();
    process.exit(0);
  }

  console.log(`✅ Se empezará a partir de la fila ${excelRowStart} de Excel (índice interno ${startIndex}).`);

  rl.question(`👉 ¿Hasta qué número de fila de EXCEL quieres procesar? (Enter = hasta la última, máximo ${totalFilasExcel}): `, answerEnd => {
    let endExclusiveIndex;

    const trimmed = answerEnd.trim();

    if (trimmed === '') {
      // Hasta el final
      endExclusiveIndex = rows.length;
      console.log(`✅ Se procesará desde la fila ${excelRowStart} hasta la última fila con datos (${totalFilasExcel}).`);
    } else {
      let excelRowEnd = parseInt(trimmed, 10);

      if (Number.isNaN(excelRowEnd)) {
        console.log('⚠️ Entrada de fila final no válida. Se usará hasta la última fila con datos.');
        endExclusiveIndex = rows.length;
        console.log(`✅ Se procesará desde la fila ${excelRowStart} hasta la última fila con datos (${totalFilasExcel}).`);
      } else {
        if (excelRowEnd < excelRowStart) {
          console.log('⚠️ La fila final es menor que la inicial. No se procesará nada.');
          rl.close();
          process.exit(0);
        }

        if (excelRowEnd > totalFilasExcel) {
          console.log(`⚠️ La fila final (${excelRowEnd}) excede el máximo (${totalFilasExcel}). Se ajustará al máximo.`);
          excelRowEnd = totalFilasExcel;
        }

        // Conversión: fila de Excel -> índice interno exclusivo
        // rows[0] = fila 2 => última fila incluida debe ser (excelRowEnd - 2)
        // endExclusive = (excelRowEnd - 2) + 1 = excelRowEnd - 1
        endExclusiveIndex = excelRowEnd - 1;

        if (endExclusiveIndex <= startIndex) {
          console.log('⚠️ El rango calculado está vacío. No hay nada que procesar.');
          rl.close();
          process.exit(0);
        }

        if (endExclusiveIndex > rows.length) {
          endExclusiveIndex = rows.length;
        }

        console.log(`✅ Se procesará desde la fila ${excelRowStart} hasta la fila ${excelRowEnd} de Excel.`);
      }
    }

    rl.close();
    procesarRango(startIndex, endExclusiveIndex);
  });
});

// ===============================
// FUNCIÓN PRINCIPAL: CREAR ARCHIVOS EN RANGO
// ===============================
function procesarRango(startIndex, endExclusiveIndex) {
  const rowsToProcess = rows.slice(startIndex, endExclusiveIndex);

  if (rowsToProcess.length === 0) {
    console.log('⚠️ No hay filas para procesar en el rango indicado.');
    return;
  }

  // Fecha actual para nombre de archivo
  const now = new Date();
  const day = String(now.getDate()).padStart(2, '0');
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const year = now.getFullYear();
  const dateStr = `${day}-${month}-${year}`; // DD-MM-AAAA

  let fileCounter = 1;

  for (let i = 0; i < rowsToProcess.length; i += CHUNK_SIZE) {
    const chunk = rowsToProcess.slice(i, i + CHUNK_SIZE);

    // Índices globales (dentro de rows) para este bloque
    const blockGlobalStartIndex = startIndex + i;
    const blockGlobalEndIndex = blockGlobalStartIndex + chunk.length - 1;

    // Convertir a filas de Excel (recordar: rows[0] = fila 2)
    const excelStartRowForBlock = blockGlobalStartIndex + 2;
    const excelEndRowForBlock = blockGlobalEndIndex + 2;

    const sheetData = [
      ['nome', 'numero', 'e-mail'], // Encabezados en portugués
      ...chunk.map(([nombre, telefono]) => [
        nombre,     // nome
        telefono,   // numero
        ''          // e-mail vacío
      ])
    ];

    const newWb = XLSX.utils.book_new();
    const newSheet = XLSX.utils.aoa_to_sheet(sheetData);

    XLSX.utils.book_append_sheet(newWb, newSheet, 'Hoja1');

    // Nombre del archivo con rango y fecha: (inicio-fin)_DD-MM-AAAA.xlsx
    const fileName = `(${excelStartRowForBlock}-${excelEndRowForBlock})_${dateStr}.xlsx`;
    const outputPath = path.join(OUTPUT_DIR, fileName);

    XLSX.writeFile(newWb, outputPath);

    console.log(`✅ Generado: ${outputPath} (${chunk.length} registros)`);
    fileCounter++;
  }

  console.log('🎉 Proceso completado con éxito.');
}
