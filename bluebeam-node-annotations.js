// read locations.xlsx and create a Bluebeam annotation file


const xlsx = require('xlsx');
const fs = require('fs');
const xmlbuilder = require('xmlbuilder');

// Read the Excel spreadsheet
const workbook = xlsx.readFile('locations.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];

// Get the range of cells in column H (location names)
const locationRange = xlsx.utils.decode_range(sheet['!ref']).e.c >= 7 ? 'H2:H' : '';
const locations = xlsx.utils.sheet_to_json(sheet, { range: locationRange });

// Create a unique RGBA color for each location
const colors = new Map();
let colorIndex = 0;
for (const location of locations) {
  if (!colors.has(location.location)) {
    colors.set(location.location, colorIndex);
    colorIndex++;
  }
}

// Create the XML document
const xml = xmlbuilder.create('Document', { version: '1.0', encoding: 'UTF-8' })
  .ele('Page', { Index: '0' })
    .ele('Label', '[1] ^EB-105-D0-A01').up()
    .ele('Width', '3370').up()
    .ele('Height', '2384').up();

// Add annotations for each location
for (const location of locations) {
  const color = colors.get(location.location);
  const layerName = location.location.replace(/\s+/g, '_');
  const layerColor = `rgba(${color}, ${color}, ${color}, 1)`;

  xml.ele('Annotation')
    .ele('Page', '[1] ^EB-105-D0-A01').up()
    .ele('Contents').up()
    .ele('ModDate', '2022-08-10T09:27:36.0000000Z').up()
    .ele('Color', layerColor).up()
    .ele('Type', 'Line').up()
    .ele('ID', 'PXLYXQROGFPUSQLX').up()
    .ele('TypeInternal', 'Bluebeam.PDF.Annotations.AnnotationLine').up()
    .ele('Raw', '789c5d8f316f83301085ff8ab7c010dfd9d8c644882189d285340413950a75a089ab525580a833e4dfd76952a9aa4ebae1bb93def7d214cce730daea325a825074eef85e9e7bc238e48d929a329548c24522692418f943a4a018b117a882afb13db56fc364db10567ebb6ee8d7adb3c17ac19173d40c31e131e21ce319e22c04737efd08f2aeb7213c6e83a2ce9feb7db97bd81407b3
