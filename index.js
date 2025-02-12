// "მომენტის" ბიბლიოთეკის იმპორტი თარიღის/დროის მანიპულირებისთვის
const moment = require("moment");

// "xlsx" ბიბლიოთეკის იმპორტი Excel ფაილების შესაქმნელად და მანიპულირებისთვის.
const XLSX = require("xlsx");

// შექმენით ახალი სამუშაო წიგნის ობიექტი XLSX-ის სასარგებლო მეთოდის გამოყენებით.
const newWorkbook = XLSX.utils.book_new();

// ცარიელი სამუშაო ფურცლის ობიექტის ინიცირება, რომელიც მოგვიანებით შეივსება უჯრედის მონაცემებით.
const newWorksheet = {};

// განსაზღვრეთ მონაცემთა მნიშვნელობების მასივი ექსპორტისთვის.
const dataForExport = [
  'მაისური', 'თანამედროვე',
  '2', '1 000', '0.01', '1','0',
  '', null, '123456789012345678901234567890', '-0', '+0',
  '-', '10,5', Number.MIN_VALUE, Number.MAX_VALUE, Infinity,
  NaN, undefined, -0, +0, -0.5, { value: '-', error: 'რიცხვითი ტიპის მნიშვნელობა არასწორადაა მითითებული' },
  '2024-03-22 13:03:03 +06:00', '2024-03-22 13:03:03 +02:00', new Date('2024-11-22T13:03:03'),
  moment('2024-11-22T13:03:03Z+06:00', 'YYYY-MM-DDTHH:mm:ss.SSSZ', `${moment().format('YYYY-MM-DD')}T08:03:03.000Z`),
  'USD', 'AED', '1210\\123456789012345678901\\002', '1210\\'
];

// Excel-ის სვეტების კონფიგურაციის განსაზღვრა.
// მასივის თითოეული ობიექტში განსაზღვრავს სვეტის მდებარეობა (მაგ., 'A'),
// უჯრედის ტიპი ('s' სტრიქონისთვის, 'n' რიცხვისთვის, 'd' თარიღისთვის, 'b' ლოგინისთვის, 'e' შეცდომისთვის),
// და სათაური (ქართულად).
const columns = [
  { column: 'A', type: 's', title: 'სტრიქონი' },  
  { column: 'B', type: 'n', title: 'რიცხვი' },  
  { column: 'C', type: 'd', title: 'თარიღი' },  
  { column: 'D', type: 'b', title: 'ლოგიკური' },   
  { column: 'E', type: 'e', title: 'შეცდომა' } 
];

// ფუქცელში სათაურები უყენდება თითოეულ სვეტს
columns.forEach(({ column, title }) => {
  const cellAddress = column + "1";
  newWorksheet[cellAddress] = { t: "s", v: title, w: title };
});

// ხდება მნიშვნელობების ჩაწერა ექსელში ყველა ტიპის მიხედვით კონკრეტულ სვეტებში
dataForExport.forEach((value, index) => {
  columns.forEach(({ column, type }) => {
    const cellAddress = column + (index + 2);
    // თარიღის ტიპზე ჭირდება დამატებითი შეზღუდვები რადგანაც თუ არ შევზღუდე არ შეიქმნება ფაილი
    if (type === "d"){
      if (value === undefined || value === null || typeof value === "number" || (typeof value === "object" && "error" in value)){
        newWorksheet[cellAddress] = { t: 's', v: 'შეცდომა: ვერ გარდაიქმნა', w: 'ვაშლი' }
      }else{
        newWorksheet[cellAddress] = { t: type, v: value, w: 'ვაშლი' };
      }
    }else{
      newWorksheet[cellAddress] = { t: type, v: value, w: 'ვაშლი' };
    }
  });
});

// სამუშაო ფურცლის დიაპაზონის დაყენება
newWorksheet["!ref"] = `A1:E${dataForExport.length + 1}`;

// შევსებული სამუშაო ფურცლის დამატება სამუშაო წიგნში ფურცლის სახელწოდებით "new_dates".
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "new_dates");

// განსაზღვრეთ ფაილის გზა, სადაც შეინახება ახალი Excel სამუშაო წიგნი.
const newFilePath = "output/new_dates.xlsx";

// ჩაწერს ექსელის ფაილს ბილიკის და სამუშაო წიგნის ობიექტის მიხედვით
XLSX.writeFile(newWorkbook, newFilePath);



