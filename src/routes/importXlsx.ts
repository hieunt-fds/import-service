import { processXLSX } from '@services/import';
import express from 'express';
import multer from 'multer';
import path from 'path';
var upload = multer({
  storage: multer.diskStorage({
    destination: function (_req, file, cb) {
      if (file.fieldname === "file") {
        cb(null, './uploads/xlsx/')
      }
      else if (file.fieldname === "tepdinhkem") {
        cb(null, './uploads/tepdinhkem/');
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
router.post('/:database/import', upload.fields([{
  name: 'file', maxCount: 1
}, {
  name: 'tepdinhkem', maxCount: 100
}]), async function (req, res) {
  if (req.files) {
    const files = req.files as { [fieldname: string]: Express.Multer.File[] };
    const metadata = await processXLSX(files, req.body.cacheDanhMuc, req.params.database, req.body.site);
    res.status(200).send(metadata)
  }
  else {
    res.status(400).send('File not found');
  }
})
export default router