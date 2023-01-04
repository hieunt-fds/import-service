import XLSX, { WorkBook, WorkSheet } from 'xlsx';
import DBUtils from '@controller/mongodb'
import { _client, _clientGridFS } from "@db/mongodb";
import { object as convertToObject } from 'dot-object'
import { getDanhMuc } from './danh_muc';
import axios from 'axios';
function addMetadataImport(record: any, fileName: string, site?: string) {
  let data = record;
  data['sourceRef'] = `${fileName}`;
  data['username'] = `ImportSevice`;
  data['openAccess'] = 0;
  data['order'] = 0;
  data['site'] = site || 'csdl_mt';
  data['storage'] = '03_import';
  data['TrangThaiDuLieu'] = {
    "_source": {
      "MaMuc": "01",
      "TenMuc": "Sơ bộ",
      "type": "C_TrangThaiDuLieu"
    }
  };
  data["accessRoles"] = [
    {
      "shortName": "admin",
      "permission": "2"
    },
    {
      "shortName": "AdminData",
      "permission": "2"
    }
  ]
  return data;
}
async function buildTepDuLieu(worksheet: WorkSheet, database: string, fileName: string, fileDinhKem?: Express.Multer.File[]) {
  if (!fileDinhKem) return;
  const sheetData: any = XLSX.utils.sheet_to_json(worksheet);
  sheetData.splice(0, 1);
  for (let index in sheetData) {
    sheetData[index]['fileName'] = `${sheetData[index]['TenTep']}.${sheetData[index]['DinhDang']}`;
    sheetData[index]['sourceRefId'] = `${fileName}___${sheetData[index][Object.keys(sheetData[index])[0]]}___${sheetData[index]['fileName']}`;
    for (let fileExpress of fileDinhKem) {
      if (fileExpress.originalname.toLocaleLowerCase() === sheetData[index].fileName.toLocaleLowerCase()) {
        let fileUploaded = await DBUtils.uploadExpressFile(_clientGridFS, "T_TepDuLieu", sheetData[index]['sourceRefId'], fileExpress);
        if (fileUploaded) {
          sheetData[index]['uploadData'] = {
            "bucketName": "T_TepDuLieu",
            "chunkSize": 102400,
            "originalname": fileUploaded.filename,
            "encoding": "7bit",
            "filename": fileUploaded.filename,
            "size": fileUploaded.chunkSizeBytes,
            "uploadDate": new Date().toISOString(),
            "id": String(fileUploaded.id),
            "contentType": fileUploaded.options.contentType || "",
          }
        }
        break;
      }
    }
    const dataToCreate = addMetadataImport(JSON.parse(JSON.stringify(sheetData[index])), fileName);
    let created = await DBUtils.createOneIfNotExist(_client, {
      dbName: database,
      collectionName: "T_TepDuLieu",
      filter: {
        sourceRefId: sheetData[index]['sourceRefId']
      },
      insertData: dataToCreate
    })
    sheetData[index]["idTepDuLieu"] = String(created.upsertedId);
  }
  return groupBy(sheetData, getHeaderRow(worksheet)[0]);
}
function groupBy(xs: any[], key: string) {
  return xs.reduce((rv, x) => {
    (rv[x[key]] = rv[x[key]] || []).push(x);
    return rv;
  }, {});
};
function getHeaderRow(worksheet: any): any {
  if (worksheet) {
    let headers = [];
    var range = XLSX.utils.decode_range(worksheet['!ref'] || '');
    var C = range.s.c;
    for (C = range.s.c; C <= range.e.c; ++C) {
      var cell = worksheet[XLSX.utils.encode_cell({ c: C, r: 0 })] /* find the cell in the first row */
      var hdr = "Cột không tên " + C; // <-- replace with your desired default 
      if (cell && cell.t) hdr = XLSX.utils.format_cell(cell);
      headers.push(hdr);
    }
    return headers;
  }
  return;
}
function findFirstColumnKey(columnName: string | undefined) {
  const regx = new RegExp(/^\w+/gi);
  return columnName?.match(regx)?.[0];
}
async function buildS_Data(worksheet: any, cacheDanhMuc: string, database: string) {
  const sheetData: any = XLSX.utils.sheet_to_json(worksheet);
  sheetData.splice(0, 1);
  const danhMucData: any = {};
  for (let index in sheetData) {
    for (let colName in sheetData[index]) {
      if (colName.startsWith("!")) {
        //Ignore column
        delete sheetData[index][colName];
        continue;
      }
      if (colName.indexOf("___") > -1) {
        let [key, ...listConfig] = colName.split("___");
        if (key.endsWith("[]")) {
          let keyToSave = key.replace("[]", "");
          let [danhMuc, keySearch, keyToADD] = listConfig;

          //default value if not exist
          keySearch = keySearch || "TenMuc";
          keyToADD = keyToADD || "MaMuc";
          let config = {
            DanhMuc: danhMuc,
            KeySearch: keySearch,
            Fields: (keyToADD || "MaMuc").split("|"),
          }
          danhMucData[danhMuc] = danhMucData[danhMuc] || await getDanhMuc(database, config, cacheDanhMuc);
          if (danhMucData[danhMuc]) {
            let lstValue = String(sheetData[index][colName]).split("||");
            let finalValue = [];
            for (let val of lstValue) {
              if (danhMucData[danhMuc][val]) {
                finalValue.push(danhMucData[danhMuc][val])
              }
              else {
                finalValue.push({
                  _source: {
                    [keySearch]: val
                  }
                })
              }
            }
            sheetData[index][keyToSave] = [...sheetData[index][keyToSave] || [], ...finalValue]
          }
          else {
            return {
              status: "error",
              msg: `${danhMuc} not found!`
            }
          }
          delete sheetData[index][colName];
          // Danh mục nhiều dữ liệu tên cột key[] dữ liệu phân cách bởi ||
          // 3. ${Tên field}___${Tên danh mục}___${Key để tìm của danh mục}___${Key kèm theo phân cách bằng '|' }
          //   Mặc định MaMuc TenMuc có thể bỏ trống ${Tên field}___${Tên danh mục}
          // danhMucData[config[sheet][column].DanhMuc] = await getDanhMuc(database, config[sheet][column], cacheDanhMuc);
        }
        else {
          // 3. ${Tên field}___${Tên danh mục}___${Key để tìm của danh mục}___${Key kèm theo phân cách bằng '|' }
          //   Mặc định MaMuc TenMuc có thể bỏ trống ${Tên field}___${Tên danh mục}
          // danhMucData[config[sheet][column].DanhMuc] = await getDanhMuc(database, config[sheet][column], cacheDanhMuc);
          let keyToSave = key.replace("[]", "");
          let [danhMuc, keySearch, keyToADD] = listConfig;
          //default value if not exist
          keySearch = keySearch || "TenMuc";
          keyToADD = keyToADD || "MaMuc";
          let config = {
            DanhMuc: danhMuc,
            KeySearch: keySearch,
            Fields: (keyToADD || "MaMuc").split("|"),
          }
          danhMucData[danhMuc] = danhMucData[danhMuc] || await getDanhMuc(database, config, cacheDanhMuc);
          if (danhMucData[danhMuc]) {
            if (danhMucData[danhMuc][sheetData[index][colName]]) {
              sheetData[index][keyToSave] = danhMucData[danhMuc][sheetData[index][colName]]
            }
            else {
              sheetData[index][keyToSave] = {
                _source: {
                  [keySearch]: sheetData[index][colName]
                }
              }
            }
            delete sheetData[index][colName];
          }
          else {
            return {
              status: "error",
              msg: `${danhMuc} not found!`
            }
          }
        }

      }
    }
    sheetData[index] = convertToObject(sheetData[index])
  }
  return groupBy(sheetData, getHeaderRow(worksheet)[0])
}
async function searchElastic(db: string, collection: string, queryEs: any, aggs: any, sort: any, includeField: string[]) {
  const host = process.env.VUEJX_URL || "http://vuejx-core:3000";
  const token = process.env.VUEJX_TOKEN;
  const query = `query search($token: String, $body: JSON, $db: String, $collection: String) {
        results: search(token: $token, body: $body, db: $db, collection: $collection )
      }`
  const queryAggs = `query search($token: String, $body: JSON, $db: String, $collection: String) {
  results: aggs(token: $token, body: $body, db: $db, collection: $collection )
}`

  let bodyQuery: any = {
    "size": aggs ? 0 : 9999,
    "query": queryEs,
    sort: sort || [{
      'createdAt': 'desc'
    }]
  };
  if (aggs) {
    bodyQuery['aggs'] = aggs;
  }
  if (includeField) {
    bodyQuery['_source'] = {
      "includes": includeField
    }
  }
  let vConfig = {
    timeout: 300000,
    maxContentLength: 524288900,
    maxBodyLength: 524288900
  };
  return await axios.post(host, {
    query: aggs ? queryAggs : query,
    variables: {
      token: token,
      db: db,
      collection: collection.toLowerCase(),
      body: bodyQuery
    }
  }, vConfig).then(function (response) {
    if (aggs) {
      if (response.data.data.results) {
        return response.data.data.results.aggregations;
      }
    }
    else if (response.data.data.results && response.data.data.results.hits) {
      return response.data.data.results.hits.hits;
    }
    else return;
  }).catch(function (xhr) {
    console.log('xhrxhr', xhr);
  })
}

async function buildT_Data(worksheet: WorkSheet, _Sdata: any, cacheDanhMuc: string, database: string, _fileData: any) {
  const sheetData: any = XLSX.utils.sheet_to_json(worksheet);
  const danhMucData: any = {};
  sheetData.splice(0, 1);

  for (let index in sheetData) {

    //sheetData[index] : 1 rowData
    for (let colName in sheetData[index]) {
      //
      // Case 1. Tên cột: "!..." => bỏ cả cột
      // 
      if (colName.startsWith("!")) {
        //Ignore column
        delete sheetData[index][colName];
        continue;
      }

      if (colName.indexOf("___") > -1) {
        // idTest*___S_ABC|S_XYZ(AAA)
        let [key, ...listConfig] = colName.split("___");

        if (key.endsWith("*")) {
          //
          // Case 2. ${fieldToCheck}*___${SheetName}(keyToSave)
          //   - *: đánh dấu cột lấy dữ liệu từ sheet khác
          //   - Tên Sheet: => phân cách bằng "|" 
          //   - Mặc định cột đầu trong các sheet con dùng để so sánh với cột fieldToCheck ở sheet cha
          //   - keyToSave: tên trường lưu thành mảng ở object sheet cha. Nếu không có sẽ lấy sheetName bỏ "S_". VD: S_HanNgachXaThai => HanNgachXaThai.
          // 
          sheetData[index][key.replace("*", "")] = sheetData[index][colName];
          if (listConfig[0].indexOf("|")) {
            for (let config of listConfig[0].split("|")) {
              let prebuildDataToGet = config; // S_ABC(XYZ)
              let keyToSave = key.replace("*", "");
              if (config.indexOf("(") > -1) {
                const filter = new RegExp(/(.+?)\((.+?)\)/gi);
                prebuildDataToGet = config.replace(filter, "$1"); // S_ABC
                keyToSave = config.replace(filter, "$2"); // XYZ
              }
              else {
                keyToSave = config.replace("S_", ""); // ABC
              }
              if (prebuildDataToGet === "T_TepDuLieu") {
                if (_fileData) {
                  sheetData[index][keyToSave] = _fileData[sheetData[index][colName]];
                }
              }
              else if (_Sdata[prebuildDataToGet]) {
                sheetData[index][keyToSave] = _Sdata[prebuildDataToGet][sheetData[index][colName]];
              }
            }
          }
          else {
            let prebuildDataToGet = listConfig[0]; // S_ABC(XYZ)
            let keyToSave = key.replace("*", "");
            if (listConfig[0].indexOf("(") > -1) {
              const filter = new RegExp(/(.+?)\((.+?)\)/gi);
              prebuildDataToGet = listConfig[0].replace(filter, "$1"); // S_ABC
              keyToSave = listConfig[0].replace(filter, "$2"); // XYZ
            }
            else {
              keyToSave = listConfig[0].replace("S_", ""); // ABC
            }
            if (prebuildDataToGet === "T_TepDuLieu") {
              if (_fileData) {
                sheetData[index][keyToSave] = _fileData[sheetData[index][colName]];
              }
            }
            else if (_Sdata[prebuildDataToGet]) {
              sheetData[index][keyToSave] = _Sdata[prebuildDataToGet][sheetData[index][colName]];
            }
          }
          // clean up
          delete sheetData[index][colName];
        }
        else if (key.endsWith("[]")) {
          let keyToSave = key.replace("[]", "");
          let [danhMuc, keySearch, keyToADD] = listConfig;

          //default value if not exist
          keySearch = keySearch || "TenMuc";
          keyToADD = keyToADD || "MaMuc";
          let config = {
            DanhMuc: danhMuc,
            KeySearch: keySearch,
            Fields: (keyToADD || "MaMuc").split("|"),
          }
          danhMucData[danhMuc] = danhMucData[danhMuc] || await getDanhMuc(database, config, cacheDanhMuc);
          if (danhMucData[danhMuc]) {
            let lstValue = String(sheetData[index][colName]).split("||");
            let finalValue = [];
            for (let val of lstValue) {
              if (danhMucData[danhMuc][val]) {
                finalValue.push(danhMucData[danhMuc][val])
              }
              else {
                finalValue.push({
                  _source: {
                    [keySearch]: val
                  }
                })
              }
            }
            sheetData[index][keyToSave] = [...sheetData[index][keyToSave] || [], ...finalValue]
          }
          else {
            return {
              status: "error",
              msg: `${danhMuc} not found!`
            }
          }
          delete sheetData[index][colName];
          // Danh mục nhiều dữ liệu tên cột key[] dữ liệu phân cách bởi ||
          // 3. ${Tên field}___${Tên danh mục}___${Key để tìm của danh mục}___${Key kèm theo phân cách bằng '|' }
          //   Mặc định MaMuc TenMuc có thể bỏ trống ${Tên field}___${Tên danh mục}
          // danhMucData[config[sheet][column].DanhMuc] = await getDanhMuc(database, config[sheet][column], cacheDanhMuc);
        }
        else {
          // 3. ${Tên field}___${Tên danh mục}___${Key để tìm của danh mục}___${Key kèm theo phân cách bằng '|' }
          //   Mặc định MaMuc TenMuc có thể bỏ trống ${Tên field}___${Tên danh mục}
          // danhMucData[config[sheet][column].DanhMuc] = await getDanhMuc(database, config[sheet][column], cacheDanhMuc);
          let keyToSave = key.replace("[]", "");
          let [danhMuc, keySearch, keyToADD] = listConfig;
          //default value if not exist
          keySearch = keySearch || "TenMuc";
          keyToADD = keyToADD || "MaMuc";
          let config = {
            DanhMuc: danhMuc,
            KeySearch: keySearch,
            Fields: (keyToADD || "MaMuc").split("|"),
          }
          danhMucData[danhMuc] = danhMucData[danhMuc] || await getDanhMuc(database, config, cacheDanhMuc);
          if (danhMucData[danhMuc]) {
            if (danhMucData[danhMuc][sheetData[index][colName]]) {
              sheetData[index][keyToSave] = danhMucData[danhMuc][sheetData[index][colName]]
            }
            else {
              sheetData[index][keyToSave] = {
                _source: {
                  [keySearch]: sheetData[index][colName]
                }
              }
            }
            delete sheetData[index][colName];
          }
          else {
            return {
              status: "error",
              msg: `${danhMuc} not found!`
            }
          }
        }
      }
      else {
        // normal key text
        if (colName.endsWith("[]")) {
          let keyToSave = colName.replace("[]", "");
          sheetData[index][keyToSave] = [...sheetData[index][keyToSave] || [], ...String(sheetData[index][colName]).split("||")]
          delete sheetData[index][colName];
        }
      }
    }
  }
  return sheetData;
}

async function bulkCreateDB(lstData: any[], database: string, collection: string, worksheet: WorkBook, fileName: string, site?: string) {
  const bulkService = await DBUtils.bulkCreateOneIfNotExist(_client, {
    dbName: database,
    collectionName: collection
  })
  let ViTriBanGhiImport = 3
  for (let record of lstData) {
    let dataToCreate: any = {};
    dataToCreate['type'] = collection;
    dataToCreate['ViTriBanGhiImport'] = ViTriBanGhiImport
    dataToCreate = {
      ...dataToCreate,
      ...addMetadataImport(record, fileName, site),
      ...database == 'CSDL_MTQA' ? {
        "TrangThaiDuLieu": {
          "_source": {
            "MaMuc": "01",
            "TenMuc": "Sơ bộ",
            "type": "C_TrangThaiDuLieu"
          }
        },
        "storage": collection.startsWith('C_') ? 'regular' : "03_import"
      } : {},
    };
    await bulkService.bulkUpsertAdd({
      sourceRefId: dataToCreate['sourceRef'] + "___" + record[findFirstColumnKey(getHeaderRow(worksheet.Sheets[collection])[0]) || Object.keys(record)[0]],
      "storage": collection.startsWith('C_') ? 'regular' : "03_import"
    }, dataToCreate);
    ViTriBanGhiImport++
  }
  let response;
  try {
    response = await bulkService.bulk.execute();
  }
  catch (err: any) {
    response = err.message
  }
  return response;
}

async function bulkCreateDBS_TMP(lstData: any, database: string, collection: string, worksheet: WorkBook, fileName: string) {
  console.log(findFirstColumnKey(getHeaderRow(worksheet.Sheets[collection])[0]) || Object.keys({})[0]);
  await _client.db(database).collection(collection).deleteMany({
    sourceRef: `${fileName}`,
    'storage': collection.startsWith('C_') ? 'regular' : '03_import'
  })
  const bulkService = await DBUtils.bulkCreateOneIfNotExist(_client, {
    dbName: database,
    collectionName: collection
  })

  let indexS = 0;
  if (['S_CapPhepXaNuocThai', 'S_CapPhepXaKhiThai', 'S_CapPhepTiengOnDoRung'].indexOf(collection) != -1) {
    let lstDataArray: any = [];
    for (let key in lstData) {
      for (let record of lstData[key]) {
        lstDataArray.push(record)
        let dataToCreate: any = {};
        dataToCreate['type'] = collection;
        dataToCreate = {
          ...dataToCreate,
          ...addMetadataImport(record, fileName),
          ...database == 'CSDL_MTQA' ? {
            "TrangThaiDuLieu": {
              "_source": {
                "MaMuc": "01",
                "TenMuc": "Sơ bộ",
                "type": "C_TrangThaiDuLieu"
              }
            },
            "storage": collection.startsWith('C_') ? 'regular' : "03_import"
          } : {},
        };
        await bulkService.bulkUpsertAdd({
          sourceRefId: dataToCreate['sourceRef'] + "___" + indexS,
          "storage": collection.startsWith('C_') ? 'regular' : "03_import"
        }, dataToCreate);
        indexS++
      }
    }
  } else {

    for (let key in lstData) {
      let record: any = {
        refKey: key,
        data: lstData[key]
      }
      let dataToCreate: any = {};
      dataToCreate['type'] = collection;
      dataToCreate = {
        ...dataToCreate,
        ...addMetadataImport(record, fileName),
        ...database == 'CSDL_MTQA' ? {
          "TrangThaiDuLieu": {
            "_source": {
              "MaMuc": "01",
              "TenMuc": "Sơ bộ",
              "type": "C_TrangThaiDuLieu"
            }
          },
          "storage": "03_import"
        } : {},
      };
      await bulkService.bulkUpsertAdd({
        sourceRefId: dataToCreate['sourceRef'] + "___" + indexS,
        "storage": collection.startsWith('C_') ? 'regular' : "03_import"
      }, dataToCreate);
      indexS++
    }

  }
  let response;
  try {
    response = await bulkService.bulk.execute();
  }
  catch (err: any) {
    response = err.message
  }
  return response;
}

export { buildTepDuLieu, addMetadataImport, findFirstColumnKey, searchElastic, buildS_Data, buildT_Data, bulkCreateDB, bulkCreateDBS_TMP }
