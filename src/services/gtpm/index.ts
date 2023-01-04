import { readFile } from 'fs-extra';
import XLSX from 'xlsx';
import { genDocumentToBuffer } from './docxtemplater';

async function processXLSX(files: { [fieldname: string]: Express.Multer.File[] }, sheetNo: string) {
  let xlsxBuffer = await readFile(files.file[0].path)
  var workbook = XLSX.read(xlsxBuffer, { type: "buffer" });
  let sheetData = await processSheet(workbook, parseInt(sheetNo || "0"));
  return sheetData;
}

async function processSheet(workbook: XLSX.WorkBook, sheetNo: number) {
  // var sheetNo = parseInt(req.body ? req.body.sheetNo : "1");
  // var workbook = XLSX.readFile('/app/uploads/'+ req.file.filename);
  var first_sheet = workbook.SheetNames[(sheetNo || 4) - 1]; //sheet 4 excel
  var worksheet = workbook.Sheets[first_sheet];
  var range = XLSX.utils.decode_range(worksheet['!ref'] || "");

  // row loop
  var sheetObj: any = [];
  let index = -1;
  for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
    // Example: Get second cell in each row, i.e. Column "B"
    const cotA = worksheet[XLSX.utils.encode_cell({ r: rowNum, c: 0 })];
    const cotB = worksheet[XLSX.utils.encode_cell({ r: rowNum, c: 1 })];
    const cotC = worksheet[XLSX.utils.encode_cell({ r: rowNum, c: 2 })];
    const cotD = worksheet[XLSX.utils.encode_cell({ r: rowNum, c: 3 })];
    const cotE = worksheet[XLSX.utils.encode_cell({ r: rowNum, c: 4 })];
    const cotF = worksheet[XLSX.utils.encode_cell({ r: rowNum, c: 5 })];

    if (cotA) {
      if (parseInt(cotA.v)) {
        sheetObj.push({})
        index++;
        if (parseInt(cotA.v) === 1) {
          let aboveA = worksheet[XLSX.utils.encode_cell({ r: rowNum - 1, c: 0 })];
          let aboveB = worksheet[XLSX.utils.encode_cell({ r: rowNum - 1, c: 1 })];
          sheetObj[index].SoMuc = aboveA ? aboveA.v : '';
          sheetObj[index].TenMuc = aboveB ? aboveB.v : '';

        }
        sheetObj[index].STT = cotA ? cotA.v : '';
        sheetObj[index].TenUC = cotB ? cotB.v : '';
        sheetObj[index].TacNhanChinh = cotC ? cotC.v : '';
        sheetObj[index].TacNhanPhu = cotD ? cotD.v : 'Không có';
        sheetObj[index].BMT = cotF ? cotF.v : '';
      }
    } else {
      if (!sheetObj[index].MoTaUC) {
        sheetObj[index].MoTaUC = []
        sheetObj[index].MoTaUC.push(cotE ? cotE.v : '',)
      } else {
        sheetObj[index].MoTaUC.push(cotE ? cotE.v : '',)
      }
    }
  }
  return await genDocumentToBuffer(sheetObj)
}

export {
  processXLSX
}