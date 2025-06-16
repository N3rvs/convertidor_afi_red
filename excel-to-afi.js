const fs = require('fs');
const XLSX = require('xlsx');

function usage() {
  console.log('Usage: node excel-to-afi.js <input.xlsx> <output.afi>');
  process.exit(1);
}

function pad(value, length) {
  const str = (value ?? '').toString();
  return str.length >= length ? str.slice(0, length) : str.padEnd(length, ' ');
}

const [, , input, output] = process.argv;
if (!input || !output) {
  usage();
}

const workbook = XLSX.readFile(input);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
const dataRows = rows.slice(1); // ignore header row

function buildHeader() {
  const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  return pad('01' + pad('SILTRA', 8) + pad(date, 8), 160);
}

function buildRecord(row) {
  const nss = pad(row[0], 9);
  const dni = pad(row[1], 9);
  const firstName = pad(row[2], 30);
  const lastName1 = pad(row[3], 30);
  const lastName2 = pad(row[4], 30);
  const company = pad(row[5], 20);
  const startDate = pad(row[6], 10);
  return pad('02' + nss + dni + firstName + lastName1 + lastName2 + company + startDate, 160);
}

function buildTrailer(count) {
  return pad('99' + pad(count.toString(), 6), 160);
}

const lines = [buildHeader(), ...dataRows.map(buildRecord), buildTrailer(dataRows.length)];
fs.writeFileSync(output, lines.join('\r\n'), { encoding: 'latin1' });
console.log(`AFI file written to ${output}`);
