
const PizZip = require("pizzip");
import Docxtemplater from "docxtemplater";
const fs = require("fs-extra");
const expressionParser = require("docxtemplater/expressions");

expressionParser.filters.default = function (input: any) {
  // Make sure that if your input is undefined, your
  // output will be undefined as well and will not
  // throw an error
  if (!input) return "";
  return input;
};

async function genDocumentToBuffer(data: any, templateFileName: string) {
  const content = fs.readFileSync(`./template/${templateFileName}`,
    "binary"
  );
  const zip = new PizZip(content);

  const docxData = {
    value: data
  }
  const doc = new Docxtemplater(zip, {
    // paragraphLoop: true,
    linebreaks: true,
    nullGetter: () => {
      return ""
    },
    parser: expressionParser
  });
  doc.setData(docxData)
  doc.render();
  const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });
  return buf
}

export { genDocumentToBuffer }