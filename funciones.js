
"use strict";

const $ = (id) => document.getElementById(id);

const fileInput = $("xmlFile");
const xmlText = $("xmlText");
const fileStatus = $("fileStatus");
const message = $("message");
const convertButton = $("convertButton");
const progress = $("progress");
const dropzone = $("dropzone");

const SHEETS = [
  { key: "compras", record: "detalleCompras", title: "Compras" },
  { key: "ventas", record: "detalleVentas", title: "Ventas" },
  { key: "ventasEstablecimiento", record: "ventaEst", title: "Ventas establecimiento" }
];

const GENERAL_LABELS = {
  TipoIDInformante: "Tipo ID Informante",
  IdInformante: "ID Informante",
  razonSocial: "Razón social",
  Anio: "Año",
  Mes: "Mes",
  numEstabRuc: "Número establecimientos RUC",
  totalVentas: "Total ventas",
  codigoOperativo: "Código operativo"
};

// Anchos de las columnas del Excel de referencia.
const WIDTHS = {
  "Datos generales": [32, 40],
  "Compras": [13,12,15,17,12,15,17,14,12,14,45,14,15,13,12,12,12,14,14,16,14,19,15,18,25,27,31,33,25,24,26,27,30,26,17,18,15,15,14],
  "Ventas": [13,15,14,17,13,20,14,15,13,12,12,13,15,24],
  "Ventas establecimiento": [12,13,12]
};


// COLUMNAS DE TEXTO Y CENTRADO
const TEXT_COLUMNS = {
  compras: ["A", "B", "C", "D", "E", "G", "H", "I", "K", "Y", "Z", "AA", "AB", "AC", "AI", "AJ", "AK", "AL"],
  ventas: ["A", "B", "C", "D", "E"],
  ventasEstablecimiento: ["A"]
};


// COLUMNAS DE CONTABILIDAD
const ACCOUNTING_COLUMNS = {
  compras: ["L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "AF", "AG", "AH"],
  ventas: ["G", "H", "I", "J", "K", "L", "M"],
  ventasEstablecimiento: ["B", "C"]
};


// COLUMNAS DE FECHAS
const DATE_COLUMNS = {
  compras: ["J", "F", "AM"],
  ventas: [],
  ventasEstablecimiento: []
};


// COLUMNAS QUE SOLAMENTE SE CENTRAN
const CENTER_COLUMNS = {
  compras: ["AD", "AE"],
  ventas: ["F", "N"],
  ventasEstablecimiento: []
};


// APLICAR FORMATOS A LAS COLUMNAS
function applyColumnFormats(ws, sectionKey) {

  // FORMATO TEXTO Y CENTRADO
  const textColumns = TEXT_COLUMNS[sectionKey] || [];

  textColumns.forEach(letter => {
    const column = ws.getColumn(letter);

    column.numFmt = "@";
    column.alignment = {
      horizontal: "center",
      vertical: "middle"
    };
  });


  // FORMATO NUMÉRICO Y ALINEACIÓN A LA DERECHA
  const accountingColumns = ACCOUNTING_COLUMNS[sectionKey] || [];

  accountingColumns.forEach(letter => {
    const column = ws.getColumn(letter);

    column.numFmt = "#,##0.00";
    column.alignment = {
      horizontal: "right",
      vertical: "middle"
    };

    // Convierte los valores de texto a números.
    for (let row = 2; row <= ws.rowCount; row++) {
      const cell = ws.getCell(`${letter}${row}`);

      if (cell.value !== null && cell.value !== "") {
        const number = Number(String(cell.value).trim());

        if (Number.isFinite(number)) {
          cell.value = number;
        }
      }
    }
  });


  // SOLAMENTE CENTRADO
  const centerColumns = CENTER_COLUMNS[sectionKey] || [];

  centerColumns.forEach(letter => {
    ws.getColumn(letter).alignment = {
      horizontal: "center",
      vertical: "middle"
    };
  });


  // FORMATO FECHA Y CENTRADO
  const dateColumns = DATE_COLUMNS[sectionKey] || [];

  dateColumns.forEach(letter => {
    const column = ws.getColumn(letter);

    column.numFmt = "dd/mm/yyyy";
    column.alignment = {
      horizontal: "center",
      vertical: "middle"
    };

    for (let row = 2; row <= ws.rowCount; row++) {

      const cell = ws.getCell(`${letter}${row}`);

      if (cell.value === null || cell.value === "") continue;

      // El XML utiliza DD/MM/YYYY.
      const parts = String(cell.value).trim().split("/");

      if (parts.length === 3) {
        const day = Number(parts[0]);
        const month = Number(parts[1]);
        const year = Number(parts[2]);

        const date = new Date(Date.UTC(year, month - 1, day));

        // Comprueba que la fecha sea válida.
        if (
          date.getUTCDate() === day &&
          date.getUTCMonth() === month - 1 &&
          date.getUTCFullYear() === year
        ) {
          cell.value = date;
        }
      }
    }
  });

}


// NOMBRE LOCAL DE LOS ELEMENTOS XML
function localName(node) {
  return node.localName || node.nodeName.split(":").pop();
}


// CONVERTIR ESTRUCTURAS ANIDADAS DEL XML
function flattenElement(element, prefix = "", output = {}) {

  const children = Array.from(element.children || []);

  if (!children.length) {
    const key = prefix || localName(element);
    output[key] = (element.textContent || "").trim();
    return output;
  }

  for (const child of children) {
    const name = prefix
      ? `${prefix}_${localName(child)}`
      : localName(child);

    flattenElement(child, name, output);
  }

  return output;
}


// VALIDAR Y LEER XML
function parseXml(xmlString) {

  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlString, "application/xml");

  const parserError = doc.querySelector("parsererror");

  if (
    parserError ||
    !doc.documentElement ||
    localName(doc.documentElement) !== "iva"
  ) {
    throw new Error(
      "El texto no parece ser un XML válido del ATS del SRI. Verifica que hayas copiado el XML completo."
    );
  }

  return doc;
}


// BUSCAR ELEMENTO DIRECTO
function findDirectChild(parent, name) {
  return Array.from(parent.children).find(
    child => localName(child) === name
  ) || null;
}


// OBTENER FILAS Y ENCABEZADOS
function makeRows(sectionNode, recordName) {

  if (!sectionNode) {
    return { headers: [], rows: [] };
  }

  const records = Array.from(sectionNode.children).filter(
    node => localName(node) === recordName
  );

  const flattened = records.map(record => flattenElement(record));

  const headers = [];

  for (const record of flattened) {
    for (const key of Object.keys(record)) {
      if (!headers.includes(key)) {
        headers.push(key);
      }
    }
  }

  const rows = flattened.map(record =>
    headers.map(header => record[header] ?? "")
  );

  return { headers, rows };
}


// FORMATO DE ENCABEZADOS
function setHeader(row) {

  row.height = 30;

  row.eachCell(cell => {
    cell.font = {
      bold: true,
      color: { argb: "FFFFFFFF" },
      name: "Arial",
      size: 10
    };

    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF1F4E78" }
    };

    cell.alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true
    };

    cell.border = {
      bottom: {
        style: "thin",
        color: { argb: "FFD9E2F3" }
      }
    };
  });
}


// CONFIGURACIÓN GENERAL DE LAS HOJAS
function setSheetLayout(ws, sheetName, headers) {

  const widths = WIDTHS[sheetName] || [];

  ws.columns = headers.map((header, index) => ({
    header,
    key: `c${index}`,
    width: widths[index] ||
      Math.min(Math.max(String(header).length + 3, 14), 32)
  }));

  setHeader(ws.getRow(1));

  ws.views = [{ state: "frozen", ySplit: 1 }];

  if (headers.length && ws.rowCount > 1) {
    ws.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: ws.rowCount, column: headers.length }
    };
  }
}


// CREAR HOJA DE DATOS GENERALES
function addGeneralSheet(workbook, root) {

  const ws = workbook.addWorksheet("Datos generales");

  ws.addRow(["Campo", "Valor"]);

  for (const child of Array.from(root.children)) {

    const key = localName(child);

    if (
      ["compras", "ventas", "ventasEstablecimiento"].includes(key)
    ) {
      continue;
    }

    ws.addRow([
      GENERAL_LABELS[key] || key,
      (child.textContent || "").trim()
    ]);
  }

  setSheetLayout(ws, "Datos generales", ["Campo", "Valor"]);

  ws.getColumn(1).width = 32;
  ws.getColumn(2).width = 40;

  for (let row = 2; row <= ws.rowCount; row++) {
    ws.getCell(row, 2).numFmt = "@";
  }
}


// CREAR HOJAS DE COMPRAS, VENTAS Y ESTABLECIMIENTOS
function addDataSheet(workbook, root, section) {

  const sectionNode = findDirectChild(root, section.key);

  const { headers, rows } = makeRows(
    sectionNode,
    section.record
  );

  const ws = workbook.addWorksheet(section.title);

  if (!headers.length) {
    ws.addRow(["No se encontraron registros en esta sección"]);
    ws.getCell("A1").font = {
      bold: true,
      color: { argb: "FF1F4E78" }
    };
    ws.getColumn(1).width = 45;
    return;
  }

  ws.addRow(headers);

  for (const values of rows) {
    ws.addRow(values);
  }

  setSheetLayout(ws, section.title, headers);

  // Conserva inicialmente los valores como texto.
  for (let row = 2; row <= ws.rowCount; row++) {

    for (let col = 1; col <= headers.length; col++) {

      const cell = ws.getRow(row).getCell(col);

      if (cell.value !== null && cell.value !== undefined) {
        cell.value = String(cell.value);
      }
    }
  }

  // Aplica los formatos correspondientes.
  applyColumnFormats(ws, section.key);
}


// CONSTRUIR EL LIBRO EXCEL
function buildWorkbook(xmlString) {

  const doc = parseXml(xmlString);
  const root = doc.documentElement;

  const workbook = new ExcelJS.Workbook();

  workbook.creator = "Conversor XML ATS";
  workbook.subject = "Conversión de XML del ATS a Excel";
  workbook.created = new Date();

  addGeneralSheet(workbook, root);

  for (const section of SHEETS) {
    addDataSheet(workbook, root, section);
  }

  return workbook;
}


// NOMBRE DEL ARCHIVO EXCEL
function makeOutputName(file) {

  if (file && file.name) {
    return file.name.replace(/\.xml$/i, "") + ".xlsx";
  }

  return "ReporteATS_convertido.xlsx";
}


// PROCESAR Y DESCARGAR UN XML
async function processFile(file) {

  const content = await file.text();

  const workbook = buildWorkbook(content);

  const buffer = await workbook.xlsx.writeBuffer();

  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");

  link.href = url;
  link.download = makeOutputName(file);

  document.body.appendChild(link);
  link.click();
  link.remove();

  setTimeout(() => URL.revokeObjectURL(url), 3000);
}


// CONVERTIR UNO O VARIOS XML
async function convertAndDownload() {

  hideMessage();

  const files = Array.from(fileInput.files || []);
  const pastedContent = xmlText.value.trim();

  if (!files.length && !pastedContent) {
    showMessage(
      "Selecciona uno o varios archivos XML o pega su contenido.",
      "error"
    );
    return;
  }

  convertButton.disabled = true;
  progress.hidden = false;

  try {

    let converted = 0;

    if (files.length) {

      for (const file of files) {
        await processFile(file);
        converted++;
      }

    } else {

      const workbook = buildWorkbook(pastedContent);

      const buffer = await workbook.xlsx.writeBuffer();

      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      });

      const url = URL.createObjectURL(blob);

      const link = document.createElement("a");

      link.href = url;
      link.download = makeOutputName(null);

      document.body.appendChild(link);
      link.click();
      link.remove();

      setTimeout(() => URL.revokeObjectURL(url), 3000);

      converted++;
    }

    showMessage(
      `¡Conversión completada! Se generaron ${converted} archivos Excel.`,
      "success"
    );

  } catch (error) {

    console.error(error);

    showMessage(
      error.message || "No se pudieron generar los archivos Excel.",
      "error"
    );

  } finally {

    convertButton.disabled = false;
    progress.hidden = true;
  }
}


// MOSTRAR MENSAJES
function showMessage(text, type) {
  message.textContent = text;
  message.className = `message ${type}`;
}

function hideMessage() {
  message.textContent = "";
  message.className = "message";
}


// SELECCIONAR UNO O VARIOS ARCHIVOS
fileInput.addEventListener("change", async () => {

  const files = Array.from(fileInput.files || []);

  if (!files.length) return;

  if (files.some(file => !/\.xml$/i.test(file.name))) {
    showMessage(
      "Todos los archivos deben tener extensión .xml.",
      "error"
    );
    return;
  }

  try {

    if (files.length === 1) {

      xmlText.value = await files[0].text();

    } else {

      // No se mezclan los XML en un solo texto.
      xmlText.value = "";
    }

    fileStatus.textContent =
      `${files.length} archivo(s) cargado(s): ${files.map(file => file.name).join(", ")}`;

    hideMessage();

  } catch {

    showMessage(
      "No se pudieron leer los archivos seleccionados.",
      "error"
    );
  }
});


// BOTÓN LIMPIAR
$("clearButton").addEventListener("click", () => {

  fileInput.value = "";
  xmlText.value = "";

  fileStatus.textContent =
    "Todavía no has cargado ningún XML.";

  hideMessage();
});


// BOTÓN CONVERTIR
convertButton.addEventListener("click", convertAndDownload);


// ARRASTRAR ARCHIVOS
dropzone.addEventListener("dragover", event => {
  event.preventDefault();
  dropzone.classList.add("dragover");
});

dropzone.addEventListener("dragleave", () => {
  dropzone.classList.remove("dragover");
});

dropzone.addEventListener("drop", async event => {

  event.preventDefault();
  dropzone.classList.remove("dragover");

  const files = Array.from(event.dataTransfer.files || []);

  if (!files.length) return;

  if (files.some(file => !/\.xml$/i.test(file.name))) {
    showMessage(
      "Todos los archivos deben tener extensión .xml.",
      "error"
    );
    return;
  }

  try {

    const dataTransfer = new DataTransfer();

    for (const file of files) {
      dataTransfer.items.add(file);
    }

    fileInput.files = dataTransfer.files;

    if (files.length === 1) {

      xmlText.value = await files[0].text();

    } else {

      xmlText.value = "";
    }

    fileStatus.textContent =
      `${files.length} archivo(s) cargado(s): ${files.map(file => file.name).join(", ")}`;

    hideMessage();

  } catch {

    showMessage(
      "No se pudieron leer los archivos seleccionados.",
      "error"
    );
  }
});