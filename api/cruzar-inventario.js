import ExcelJS from "exceljs";
import formidable from "formidable";
import fs from "fs";

export const config = {
  api: {
    bodyParser: false
  }
};

// ===============================
// NORMALIZADOR UNIVERSAL
// ===============================
function normalizarTexto(texto) {
  return String(texto || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // quitar tildes
    .replace(/\s+/g, " ")
    .trim();
}

// ===============================
// DETECTOR INTELIGENTE DE COLUMNA
// ===============================
function esColumnaCodigo(valor) {
  const texto = normalizarTexto(valor);

  const palabrasClave = [
    "codigo",
    "cod",
    "codigo de barras",
    "barcode",
    "bar code",
    "id",
    "identificador",
    "identificacao",
    "identificacion",
    "serial",
    "serial number",
    "item code",
    "product code",
    "sku",
    "ean",
    "upc"
  ];

  return palabrasClave.some(palabra =>
    texto.includes(palabra)
  );
}

function encontrarColumnaCodigo(worksheet) {
  let columna = null;

  worksheet.getRow(1).eachCell((cell, col) => {
    if (esColumnaCodigo(cell.value)) {
      columna = col;
    }
  });

  return columna;
}

// ===============================
// HANDLER PRINCIPAL
// ===============================
export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Método no permitido" });
  }

  const form = formidable({ multiples: true });

  form.parse(req, async (err, fields, files) => {
    if (err) {
      return res.status(500).json({ error: err.message });
    }

    const inventarioFile = files.inventario?.[0];
    const escaneoFile = files.escaneo?.[0];

    if (!inventarioFile || !escaneoFile) {
      return res.status(400).json({ error: "Faltan archivos Excel" });
    }

    try {
      const wbInventario = new ExcelJS.Workbook();
      const wbEscaneo = new ExcelJS.Workbook();

      await wbInventario.xlsx.readFile(inventarioFile.filepath);
      await wbEscaneo.xlsx.readFile(escaneoFile.filepath);

      const wsInventario = wbInventario.worksheets[0];
      const wsEscaneo = wbEscaneo.worksheets[0];

      // ===============================
      // DETECTAR COLUMNA EN ESCANEO
      // ===============================
      const colCodigoEscaneo = encontrarColumnaCodigo(wsEscaneo);

      if (!colCodigoEscaneo) {
        return res.status(400).json({
          error: "No se encontró columna tipo código en el Excel de escaneo"
        });
      }

      // ===============================
      // EXTRAER CODIGOS ESCANEADOS
      // ===============================
      const codigosEscaneo = new Set();

      wsEscaneo.eachRow((row, i) => {
        if (i === 1) return;

        const valor = row.getCell(colCodigoEscaneo).value;
        const codigo = normalizarTexto(valor);

        if (codigo) {
          codigosEscaneo.add(codigo);
        }
      });

      // ===============================
      // DETECTAR COLUMNA EN INVENTARIO
      // ===============================
      const colCodigoInventario = encontrarColumnaCodigo(wsInventario);

      if (!colCodigoInventario) {
        return res.status(400).json({
          error: "No se encontró columna tipo código en el Excel de inventario"
        });
      }

      // ===============================
      // CRUCE Y MARCADO
      // ===============================
      let coincidencias = 0;

      wsInventario.eachRow((row, i) => {
        if (i === 1) return;

        const cell = row.getCell(colCodigoInventario);
        const codigo = normalizarTexto(cell.value);

        if (codigosEscaneo.has(codigo)) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF00FF00" }
          };
          coincidencias++;
        }
      });

      // ===============================
      // RESPUESTA FINAL
      // ===============================
      const buffer = await wbInventario.xlsx.writeBuffer();

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );

      res.setHeader(
        "Content-Disposition",
        "attachment; filename=inventario_cruzado.xlsx"
      );

      res.send(buffer);

      // Limpieza opcional
      fs.unlinkSync(inventarioFile.filepath);
      fs.unlinkSync(escaneoFile.filepath);

    } catch (error) {
      return res.status(500).json({
        error: "Error procesando archivos",
        detalle: error.message
      });
    }
  });
}
