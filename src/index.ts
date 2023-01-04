import 'dotenv/config'
import 'module-alias/register'
import bodyParser from 'body-parser';
import https from 'https';
import express from 'express';
import { ensureDir } from 'fs-extra';
import ImportXLSXRouter from "@routes/importXlsx";
import GTPMRouter from "@routes/import_gtpm";


https.globalAgent.options.rejectUnauthorized = false;

const app = express();
app.use(bodyParser.json({
  limit: "50mb"
}));
app.use(
  bodyParser.urlencoded({
    limit: "50mb",
    extended: true,
    parameterLimit: 50000,
  })
);
app.use((err: any, _req: any, res: any, _next: any) => {
  res.status(err.status || 500);
  res.json({
    message: err.message,
    error: err,
  });
});
// first run init
ensureDir("tmp/")
ensureDir("uploads/xlsx")
ensureDir("uploads/tepdinhkem")
ensureDir("uploads/gtpm")

app.use('/importXLSX', ImportXLSXRouter)
app.use('/importGTPM', GTPMRouter)
app.listen(process.env.PORT, async () => {
  console.log("Server is up!");
})
