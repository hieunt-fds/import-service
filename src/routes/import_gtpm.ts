import express from 'express';
import multer from 'multer'
import path from 'path';
import { processXLSX } from '@services/gtpm';
import { unlinkSync } from 'fs-extra';

var upload = multer({
  storage: multer.diskStorage({
    destination: function (_req, file, cb) {
      if (file.fieldname === "file") {
        cb(null, './uploads/gtpm/')
      }
    },
    filename: function (_req, file, cb) {
      cb(null, `${Date.now()}${path.extname(file.originalname)}`);
    },
  }),
  fileFilter: (_req, file, cb) => {
    file.originalname = Buffer.from(file.originalname, 'latin1').toString(
      'utf8'
    )
    cb(null, true)
  },
});

const router = express.Router();

router.get('/ping', async function (_req, res) {
  res.status(200).send("Service is up and running!")
})

router.post('/import', upload.fields([{
  name: 'file', maxCount: 1
}]), async function (req, res) {
  if (req.files) {
    const files = req.files as { [fieldname: string]: Express.Multer.File[] };
    let responseFileName = String(files?.file?.[0].originalname).replace('.xlsx', '.docx');
    const fileBuffer = await processXLSX(files, req.body?.sheetNo);
    res.writeHead(200, {
      'Content-Type': "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      'Content-disposition': 'attachment;filename=' + responseFileName,
      'Content-Length': fileBuffer.length
    });
    await unlinkSync(files?.file?.[0].path)
    res.end(fileBuffer);
  }
  else {
    res.status(400).send('File not found');
  }
})

export default router