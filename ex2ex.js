const XLSX = require('xlsx');

const readSheet = ("./oldExcel/" + process.argv[2] + ".xlsx")

const writeSheet = ("./newExcel/" + process.argv[3] + ".xlsx")

const dataToWrite = XLSX.readFile(readSheet)

XLSX.writeFile(dataToWrite, writeSheet)