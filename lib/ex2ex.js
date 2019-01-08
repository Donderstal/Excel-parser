const XLSX = require('xlsx');

const readSheet = "./oldExcel/" + process.argv[2] + ".xlsx";

const writeSheet = "./newExcel/" + process.argv[3] + ".xlsx";

const workbook = XLSX.read(readSheet);

const rawDataReport = workbook.Sheets[workbook.SheetNames[0]];

/* XLSX.writeFile(dataToWrite, writeSheet) */

/* console.log(workbook.A3.v.split('_')) */

/* console.log(XLSX.utils.encode_col(2)) */

console.log(workbook);

/* console.log(Object.keys(rawDataReport)) */