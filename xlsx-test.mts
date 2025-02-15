import moment from 'moment';
import { createRequire } from 'module';

import * as fs from 'fs';

export async function createFolderIfNotExists(folderPath: string) {
    try {
        if (!fs.existsSync(folderPath)) {
            await fs.promises.mkdir(folderPath, { recursive: true });
        }
    } catch (err) {
        console.error(`შეცდომა დირექტორიის შექმნისას: ${err}`);
    }
}

const require = createRequire(import.meta.url);

const XLSX = require('xlsx');

const wb = XLSX.utils.book_new();

const
    today = moment(),
    ws = {};

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღის ობიექტი',
    today.toDate(),
]], {dense: true});

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღის ობიექტი - cellDates: true',
    today.toDate(),
]], {dense: true, cellDates: true, origin: -1});

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღი ISO string - cellDates: true',
    today.toISOString(),
]], {dense: true, cellDates: true, origin: -1});

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღი in cell Object - cellDates: true',
    {t: 'd', v: today.toDate()},
]], {dense: true, cellDates: true, origin: -1});

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღი ISO string in cell Object - cellDates: true',
    {t: 'd', v: today.toISOString()},
]], {dense: true, cellDates: true, origin: -1});

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღი ISO string in cell Object - yyyy-mm-dd hh:mm:ss - cellDates: true UTC: false',
    {t: 'd', v: today.toISOString(), z: 'yyyy-mm-dd hh:mm:ss'},
]], {cellDates: true,  UTC: false, origin: -1});

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღი ISO string in cell Object - yyyy-mm-dd hh:mm:ss - cellDates: true UTC: true',
    {t: 'd', v: today.toISOString(), z: 'yyyy-mm-dd hh:mm:ss'},
]], {dense: true, cellDates: true, UTC: true, origin: -1});

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღი ობიექტი in cell Object - yyyy-mm-dd hh:mm:ss - cellDates: true UTC: false',
    {t: 'd', v: today.toDate(), z: 'yyyy-mm-dd hh:mm:ss'},
]], {dense: true, UTC: false, origin: -1});

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღი ობიექტი in cell Object - yyyy-mm-dd hh:mm:ss - cellDates: true UTC: true',
    {t: 'd', v: today.toDate(), z: 'yyyy-mm-dd hh:mm:ss'},
]], {dense: true, UTC: true, origin: -1});

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღი ობიექტი - cellDates: true UTC: false',
    today.toDate(),
]], {dense: true, cellDates: true, UTC: false, origin: -1});

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღი ობიექტი - cellDates: true UTC: true',
    today.toDate(),
]], {dense: true, cellDates: true, UTC: true, origin: -1});

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღი ობიექტი - UTC: false',
    today.toDate(),
]], {dense: true, UTC: false, origin: -1});

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღი ობიექტი - UTC: true',
    today.toDate(),
]], {dense: true, UTC: true, origin: -1});

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღი ობიექტი in cell Object - UTC: false',
    {t: 'd', v: today.toDate()},
]], {dense: true, UTC: false, origin: -1});

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღი ობიექტი in cell Object - UTC: true',
    {t: 'd', v: today.toDate()},
]], {dense: true, UTC: true, origin: -1});

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღი ISO string in cell Object - [hh] - cellDates: true',
    {t: 'd', v: today.toISOString(), z: '[hh]'},
]], {dense: true, cellDates: true, origin: -1});

XLSX.utils.sheet_add_aoa(ws, [[
    'თარიღის ობიექტი - dateNF: yyyy-mm-dd hh:mm:ss - cellDates: true',
    today.toDate(),
]], {dense: true, cellDates: true, dateNF: 'yyyy-mm-dd hh:mm:ss', origin: -1});

XLSX.utils.book_append_sheet(wb, ws, "test of date");
//
// console.log(ws['!data'])

await createFolderIfNotExists('build/output');
XLSX.writeFile(wb, `build/output/text-${today.valueOf()}.xlsx`);
