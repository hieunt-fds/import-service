import XLSX from 'xlsx';
// import { _client, _clientGridFS } from "@db/mongodb";
// import { object as convertToObject } from 'dot-object'
import { readFile } from 'fs-extra';
import { buildTepDuLieu, buildS_Data, buildT_Data, bulkCreateDB } from '@services/import/utils'

async function processXLSX(files: { [fieldname: string]: Express.Multer.File[] }, cacheDanhMuc: string = 'false', database: string) {
  let xlsxBuffer = await readFile(files.file[0].path)
  var workbook = XLSX.read(xlsxBuffer, { type: "buffer" });
  let sheetData = await mapConfigSheet(workbook, cacheDanhMuc, database, files.file[0].originalname, files.tepdinhkem);

  return sheetData;
}
async function mapConfigSheet(worksheet: XLSX.WorkBook, cacheDanhMuc: string = 'false', database: string, fileName: string, fileDinhKem?: Express.Multer.File[]) {
  const responseData: any = {};
  const _Sdata: any = {};
  const _Tdata: any = {};
  let _fileData: any = {};
  let lstSheet_S = worksheet.SheetNames.filter(x => x.startsWith("S_"));
  let lstSheet_T = worksheet.SheetNames.filter(x => x.startsWith("T_") && (x !== "T_TepDuLieu"));
  let lstSheet_C = worksheet.SheetNames.filter(x => x.startsWith("C_"));
  if (worksheet.Sheets["T_TepDuLieu"]) {
    _fileData = await buildTepDuLieu(worksheet.Sheets["T_TepDuLieu"], database, fileName, fileDinhKem)
  }
  for (let sheet of lstSheet_S) {
    // Build S_
    _Sdata[sheet] = await buildS_Data(worksheet.Sheets[sheet], cacheDanhMuc, database);
  }
  for (let sheet of [...lstSheet_T, ...lstSheet_C]) {
    _Tdata[sheet] = await buildT_Data(worksheet.Sheets[sheet], _Sdata, cacheDanhMuc, database, _fileData);
    if (Array.isArray(_Tdata[sheet])) {
      responseData[sheet] = await bulkCreateDB(_Tdata[sheet], database, sheet, worksheet, fileName)

    }
    else {
      responseData.err = _Tdata[sheet];
    }
  }
  return responseData
}
export {
  processXLSX
}