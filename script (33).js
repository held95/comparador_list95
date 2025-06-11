import formidable from "formidable";
import fs from "fs";
import { read, utils, writeFileXLSX } from "xlsx";

export const config = {
  api: {
    bodyParser: false,
  },
};

export default async function handler(req, res) {
  const form = formidable({ multiples: false });

  form.parse(req, async (err, fields, files) => {
    if (err || !files.arquivo_excel) {
      return res.status(400).json({ error: "Erro ao receber o arquivo." });
    }

    const filePath = files.arquivo_excel[0].filepath;

    const workbook = read(fs.readFileSync(filePath));
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = utils.sheet_to_json(sheet, { header: 1 });

    const colunaA = data.map(row => row[0]?.toString().trim()).filter(Boolean);
    const colunaB = data.map(row => row[1]?.toString().trim()).filter(Boolean);

    const normalA = colunaA.map(n => n.toLowerCase());
    const normalB = colunaB.map(n => n.toLowerCase());

    const resultado = [];

    // Parte 1: A em B
    colunaA.forEach((nomeA, i) => {
      const norm = normalA[i];
      const indexB = normalB.indexOf(norm);
      if (indexB !== -1) {
        resultado.push([nomeA, colunaB[indexB]]);
      } else {
        resultado.push([nomeA, ""]);
      }
    });

    // Parte 2: B não em A
    colunaB.forEach((nomeB, i) => {
      const norm = normalB[i];
      if (!normalA.includes(norm)) {
        resultado.push(["", nomeB]);
      }
    });

    const wb = utils.book_new();
    const ws = utils.aoa_to_sheet([["Coluna A", "Coluna B"], ...resultado]);
    utils.book_append_sheet(wb, ws, "Resultado");

    const tempPath = "/tmp/resultado_comparacao.xlsx";
    writeFileXLSX(wb, tempPath);

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=resultado_comparacao.xlsx");
    fs.createReadStream(tempPath).pipe(res);
  });
}
