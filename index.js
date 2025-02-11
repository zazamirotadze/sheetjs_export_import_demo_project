const XLSX = require("xlsx");

const filePath = 'data_source/original_dates.xlsx';

const workbook = XLSX.readFile(filePath, { cellDates: true} );
const worksheet = workbook.Sheets['original_dates'];

const newWorkbook = XLSX.utils.book_new();
const newWorksheet = {};

Object.keys(worksheet).forEach((cell) => {
  if (cell.startsWith("A") && worksheet[cell].t === "d") {
    newWorksheet[cell] = { t: "d", v: worksheet[cell].v }; 
    
  }
});

newWorksheet["!ref"] = worksheet["!ref"];

XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "new_dates");

const newFilePath = "output/new_dates.xlsx";
XLSX.writeFile(newWorkbook, newFilePath);
