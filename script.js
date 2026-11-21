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

// Anchos de las columnas del Excel de referencia. Las columnas nuevas
// que aparezcan en otros XML reciben un ancho calculado automáticamente.
const WIDTHS = {
  "Datos generales": [32, 40],
  "Compras": [13,12,15,17,12,15,17,14,12,14,45,14,15,13,12,12,12,14,14,16,14,19,15,18,25,27,31,33,25,24,26,27,30,26,17,18,15,15,14],
  "Ventas": [13,15,14,17,13,20,14,15,13,12,12,13,15,24],
  "Ventas establecimiento": [12,13,12]
};

const TEXT_COLUMNS = {
  compras: ["A", "C", "G", "H", "I", "K"],
  ventas: ["A", "B"],
  ventasEstablecimiento: ["A"]
};

function localName(node) {
  return node.localName || node.nodeName.split(":").pop();
}

function flattenElement(element, prefix = "", output = {}) {
  // Convierte estructuras anidadas del ATS en columnas con nombres unidos
  // por guion bajo, por ejemplo pagoExterior_pagoLocExt.
  const children = Array.from(element.children || []);
  if (!children.length) {
    const key = prefix || localName(element);
    output[key] = (element.textContent || "").trim();
    return output;
  }

  for (const child of children) {
    const name = prefix ? `${prefix}_${localName(child)}` : localName(child);
    flattenElement(child, name, output);
  }
  return output;
}

function parseXml(xmlString) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlString, "application/xml");
  const parserError = doc.querySelector("parsererror");
  if (parserError || !doc.documentElement || localName(doc.documentElement) !== "iva") {
    throw new Error("El texto no parece ser un XML válido del ATS del SRI. Verifica que hayas copiado el XML completo.");
  }
  return doc;
}

function findDirectChild(parent, name) {
  return Array.from(parent.children).find(child => localName(child) === name) || null;
}

function makeRows(sectionNode, recordName) {
  if (!sectionNode) return { headers: [], rows: [] };
  const records = Array.from(sectionNode.children).filter(node => localName(node) === recordName);
  const flattened = records.map(record => flattenElement(record));
  const headers = [];
  for (const record of flattened) {
    for (const key of Object.keys(record)) {
      if (!headers.includes(key)) headers.push(key);
    }
  }
  const rows = flattened.map(record => headers.map(header => record[header] ?? ""));
  return { headers, rows };
}

function setHeader(row) {
  row.height = 30;
  row.eachCell(cell => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" }, name: "Arial", size: 10 };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = { bottom: { style: "thin", color: { argb: "FFD9E2F3" } } };
  });
}

function setSheetLayout(ws, sheetName, headers, sectionKey = null) {
  const widths = WIDTHS[sheetName] || [];
  ws.columns = headers.map((header, index) => ({
    header,
    key: `c${index}`,
    width: widths[index] || Math.min(Math.max(String(header).length + 3, 14), 32)
  }));

  setHeader(ws.getRow(1));
  ws.views = [{ state: "frozen", ySplit: 1 }];
  if (headers.length && ws.rowCount > 1) {
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: ws.rowCount, column: headers.length } };
  }

  const textLetters = sectionKey ? (TEXT_COLUMNS[sectionKey] || []) : [];
  for (const letter of textLetters) {
    const columnNumber = letter.charCodeAt(0) - 64;
    for (let row = 2; row <= ws.rowCount; row++) {
      ws.getRow(row).getCell(columnNumber).numFmt = "@";
    }
  }
}

function addGeneralSheet(workbook, root) {
  const ws = workbook.addWorksheet("Datos generales");
  ws.addRow(["Campo", "Valor"]);
  for (const child of Array.from(root.children)) {
    const key = localName(child);
    if (["compras", "ventas", "ventasEstablecimiento"].includes(key)) continue;
    ws.addRow([GENERAL_LABELS[key] || key, (child.textContent || "").trim()]);
  }
  setSheetLayout(ws, "Datos generales", ["Campo", "Valor"]);
  ws.getColumn(1).width = 32;
  ws.getColumn(2).width = 40;
  for (let row = 2; row <= ws.rowCount; row++) {
    ws.getCell(row, 2).numFmt = "@";
  }
}

function addDataSheet(workbook, root, section) {
  const sectionNode = findDirectChild(root, section.key);
  const { headers, rows } = makeRows(sectionNode, section.record);
  const ws = workbook.addWorksheet(section.title);

  if (!headers.length) {
    ws.addRow(["No se encontraron registros en esta sección"]);
    ws.getCell("A1").font = { bold: true, color: { argb: "FF1F4E78" } };
    ws.getColumn(1).width = 45;
    return;
  }

  ws.addRow(headers);
  for (const values of rows) ws.addRow(values);
  setSheetLayout(ws, section.title, headers, section.key);

  // El Excel de referencia conserva los valores de datos como texto.
  // Se aplica formato de texto a las columnas de identificadores indicadas.
  for (let row = 2; row <= ws.rowCount; row++) {
    for (let col = 1; col <= headers.length; col++) {
      const cell = ws.getRow(row).getCell(col);
      if (cell.value !== null && cell.value !== undefined) cell.value = String(cell.value);
    }
  }
}

function buildWorkbook(xmlString) {
  const doc = parseXml(xmlString);
  const root = doc.documentElement;
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Conversor XML ATS";
  workbook.subject = "Conversión de XML del ATS a Excel";
  workbook.created = new Date();

  addGeneralSheet(workbook, root);
  for (const section of SHEETS) addDataSheet(workbook, root, section);
  return workbook;
}

async function convertAndDownload() {
  hideMessage();
  const content = xmlText.value.trim();
  if (!content) {
    showMessage("Selecciona un archivo XML o pega primero su contenido.", "error");
    return;
  }

  convertButton.disabled = true;
  progress.hidden = false;
  try {
    const workbook = buildWorkbook(content);
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = makeOutputName();
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);

    const doc = parseXml(content);
    const counts = SHEETS.map(section => {
      const node = findDirectChild(doc.documentElement, section.key);
      return `${section.title}: ${node ? node.children.length : 0}`;
    }).join(" · ");
    showMessage(`¡Excel generado! ${counts}. Revisa tus descargas.`, "success");
  } catch (error) {
    console.error(error);
    showMessage(error.message || "No se pudo generar el archivo Excel.", "error");
  } finally {
    convertButton.disabled = false;
    progress.hidden = true;
  }
}

function makeOutputName() {
  const selected = fileInput.files && fileInput.files[0];
  if (selected && selected.name) return selected.name.replace(/\.xml$/i, "") + ".xlsx";
  return "ReporteATS_convertido.xlsx";
}

function showMessage(text, type) {
  message.textContent = text;
  message.className = `message ${type}`;
}
function hideMessage() {
  message.textContent = "";
  message.className = "message";
}

fileInput.addEventListener("change", async () => {
  const file = fileInput.files && fileInput.files[0];
  if (!file) return;
  try {
    xmlText.value = await file.text();
    fileStatus.textContent = `Archivo cargado: ${file.name} (${Math.ceil(file.size / 1024)} KB)`;
    hideMessage();
  } catch {
    showMessage("No se pudo leer el archivo. Intenta abrirlo y pegar su contenido.", "error");
  }
});

$("clearButton").addEventListener("click", () => {
  fileInput.value = "";
  xmlText.value = "";
  fileStatus.textContent = "Todavía no has cargado ningún XML.";
  hideMessage();
});

convertButton.addEventListener("click", convertAndDownload);

dropzone.addEventListener("dragover", event => {
  event.preventDefault();
  dropzone.classList.add("dragover");
});
dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
dropzone.addEventListener("drop", async event => {
  event.preventDefault();
  dropzone.classList.remove("dragover");
  const file = event.dataTransfer.files && event.dataTransfer.files[0];
  if (!file) return;
  if (!/\.xml$/i.test(file.name) && !/xml/i.test(file.type)) {
    showMessage("Selecciona un archivo con extensión .xml.", "error");
    return;
  }
  fileInput.files = event.dataTransfer.files;
  xmlText.value = await file.text();
  fileStatus.textContent = `Archivo cargado: ${file.name} (${Math.ceil(file.size / 1024)} KB)`;
  hideMessage();
});
