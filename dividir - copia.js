const XLSX = require('xlsx');

// ===============================
// CONFIGURACIÓN
// ===============================
const INPUT_FILE = 'clientes.xlsx';   // Archivo original
const CHUNK_SIZE = 50;                // Cantidad de registros por archivo
const SHEET_INDEX = 0;                // Primera hoja del Excel

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

// ===============================
// CREAR ARCHIVOS DE 50 EN 50
// ===============================
let fileCounter = 1;

for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
  const chunk = rows.slice(i, i + CHUNK_SIZE);

  const sheetData = [
    ['nome', 'numero', 'e-mail'], // Encabezados en portugués
    ...chunk.map(([nombre, telefono]) => [
      nombre,     // nome
      telefono,   // numero
      ''           // e-mail vacío
    ])
  ];

  const newWb = XLSX.utils.book_new();
  const newSheet = XLSX.utils.aoa_to_sheet(sheetData);

  XLSX.utils.book_append_sheet(newWb, newSheet, 'Hoja1');

  const outputName = `bloque_${fileCounter}.xlsx`;
  XLSX.writeFile(newWb, outputName);

  console.log(`✅ Generado: ${outputName} (${chunk.length} registros)`);
  fileCounter++;
}

console.log('✅ Proceso completado con éxito.');
