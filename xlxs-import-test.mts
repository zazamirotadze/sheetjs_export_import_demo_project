import XLSX from 'xlsx';
import * as fs from 'fs';

/**
 * სპეციფიკურობას ანიჭებს XLSX ბიბლიოთეკას Node.js ფაილური სისტემის მოდულის გამოსაყენებლად.
 * ეს საჭიროა XLSX-ისთვის ფაილის ოპერაციების შესასრულებლად (მაგ. ფაილების წაკითხვა ან ჩაწერა)
 * Node.js გარემოში გაშვებისას.
 */
XLSX.set_fs(fs);

/**
 * ტიპი წარმოადგენს Excel-ის უჯრის მნიშვნელობას.
 *ის შეიძლება იყოს სტრიქონი, რიცხვი, ლოგიკური, ნულოვანი ან თარიღი.
 * @typedef {string | number | boolean | null | Date} ExcelRecord
 */
type ExcelRecord = string | number | boolean | null | Date;

/**
 * წარმოადგენს Excel ფაილიდან ამოღებულ ცხრილს.
 *
 * @typedef {Object} Table
 * @property {string[]} headers - სათაურის მნიშვნელობები.
 * @property {ExcelRecord[][]} records - ჩანაწერების მნიშვნელობების ორ განზომილებიანი მასივი.
 */
interface Table {
  headers: string[];
  records: ExcelRecord[][];
} 

const 
    filePath: string = 'excel_data/original_dates.xlsx',

    workbook: XLSX.WorkBook = XLSX.readFile(filePath, { cellDates: true} ),
    denseWorkbook: XLSX.WorkBook = XLSX.readFile(filePath, { dense: true, cellDates: true }),
    
    worksheet: XLSX.WorkSheet = workbook.Sheets["original_dates"],
    denseWorksheet: XLSX.WorkSheet = denseWorkbook.Sheets["original_dates"];

/**
 * @function presentWorksheetWithSheet_to_json
 * აანალიზებს Excel-ის სამუშაო ფურცელს JSON-ის მსგავს სტრუქტურაში, ამოიღებს სათაურებსა და ჩანაწერებს.
 * 
 * @param {XLSX.WorkSheet} ws - სამუშაო ფურცლის გასაანალიზებელი ობიექტი.
 * 
 * @returns {Table | undefined} - აბრუნებს ან ცხრილს ან განუსაზღვრელ მნიშვნელობას.
 */
function presentWorksheetWithSheet_to_json(ws: XLSX.WorkSheet):  Table | undefined {
    const data: Table['records'] | undefined = XLSX.utils.sheet_to_json(ws, { header: 1, UTC: true});
    return dataFromater(data);
}

/**
 * @function presentDenseWorksheet
 * აანალიზებს Excel-ის სამუშაო ფურცელს ჩართული მონაცემთა მკვრივი რეჟიმით, ამოიღებს სათაურებსა და ჩანაწერებს.
 *
 * @param {XLSX.WorkSheet} ws - სამუშაო ფურცლის ობიექტი, რომელიც უნდა გაანალიზდეს.
 *
 * @returns {Table | undefined} - - აბრუნებს ან ცხრილს ან განუსაზღვრელ მნიშვნელობას.
 */
function presentDenseWorksheet(ws: XLSX.WorkSheet): Table | undefined {
    const data: Table['records'] | undefined = ws["!data"]?.map(row => row.map(cell => cell?.v ?? null));
    return dataFromater(data);
}  

/**
 * @function presentWorksheetWithRef
 * აანალიზებს Excel-ის სამუშაო ფურცელს უჯრედის მითითების დიაპაზონის გამოყენებით.
 *
 * @param {XLSX.WorkSheet} ws - სამუშაო ფურცლის ობიექტი, რომელიც უნდა გაანალიზდეს.
 *
 * @returns {Table | undefined} - აბრუნებს ან ცხრილს ან განუსაზღვრელ მნიშვნელობას.
 */
function presentWorksheetWithRef(ws: XLSX.WorkSheet): Table | undefined {
    const ref: string | undefined = ws["!ref"];

    if (!ref) {
      return;
    }

    const data: Table['records'] = []
    const range: XLSX.Range = XLSX.utils.decode_range(ref);

    for (let row = range.s.r; row <= range.e.r; row++) {
      const rowData: ExcelRecord[] = [];
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = ws[cellAddress];
        rowData.push(cell ? cell.v : null);
      }
      if (rowData.length > 0){
        data.push(rowData)
      }
    }

    return dataFromater(data);
}

/**
 * @function dataFromater
 * დაამუშავებს მონაცემებს და დააფორმატებს ცხრილად
 *
 * @param {Table['records'] | undefined}  data - მონაცემები ფურცლიდან სტრიქონებისა და სვეტების სახით.
 *
 * @returns {Table | undefined} - აბრუნებს ან ცხრილს ან განუსაზღვრელ მნიშვნელობას.
 */
function dataFromater(data: Table['records'] | undefined): Table | undefined {
    if (!data || data.length === 0) {
        return;
    }

    const [headers, ...records]: Table['records'] = data;
    const headerStrings: Table['headers'] = headers.map(header => header?.toString() ?? '')

    return {headers: headerStrings, records};
} 

console.log(presentWorksheetWithSheet_to_json(worksheet), 'presentWorksheetWithSheet_to_json(worksheet)');
console.log(presentDenseWorksheet(denseWorksheet), 'presentDenseWorksheet(denseWorksheet)');
console.log(presentWorksheetWithRef(worksheet), 'presentWorksheetWithRef(worksheet)');