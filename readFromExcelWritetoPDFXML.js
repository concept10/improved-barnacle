const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const { createWorker } = require("tesseract.js");

async function readExcelText(excelFile) {
  const workbook = XLSX.read(excelFile, { type: "buffer" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  let text = "";
  for (let cell in sheet) {
    if (cell[0] === "A") {
      text += sheet[cell].v + "\n";
    }
  }
  return text;
}

async function readPdfText(pdfFile) {
  const worker = createWorker();
  await worker.load();
  await worker.loadLanguage("eng");
  await worker.initialize("eng");
  const { data: { text, boxes } } = await worker.recognize(pdfFile);
  await worker.terminate();
  return { text, boxes };
}

const workingDirectory = "./pdfs";

fs.readFile("sample.xlsx", async (err, excelData) => {
  if (err) return console.error(err);
  const textToFind = await readExcelText(excelData);
  fs.readdir(workingDirectory, async (err, files) => {
    if (err) return console.error(err);
    for (let file of files) {
      if (path.extname(file) === ".pdf") {
        fs.readFile(path.join(workingDirectory, file), async (err, pdfData) => {
          if (err) return console.error(err);
          const { text, boxes } = await readPdfText(pdfData);
          let xml = "<locations>\n";
          let startIndex = 0;
          while (true) {
            startIndex = text.indexOf(textToFind, startIndex);
            if (startIndex === -1) break;
            let endIndex = startIndex + textToFind.length;
            let found = false;
            for (let box of boxes) {
              if (box.text.includes(textToFind) &&
                  box.left <= startIndex && startIndex <= box.right &&
                  box.top <= endIndex && endIndex <= box.bottom) {
                xml += `  <location x="${box.left}" y="${box.top}" width="${box.width}" height="${box.height}" />\n`;
                found = true;
                break;
              }
            }
            if (!found) {
              xml += `  <location x="0" y="0" width="0" height="0" />\n`;
            }
            startIndex = endIndex;
          }
          xml += "</locations>\n";
          fs.writeFile(path.join(workingDirectory, path.basename(file, ".pdf") + ".xml"), xml, (err) => {
            if (err) return console.error(err);
            console.log(`XML file for ` + file + ` written successfully`);
