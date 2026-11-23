
"use strict";

// ======================================================
// ELEMENTOS HTML
// ======================================================

const $ = (id) => document.getElementById(id);

const fileInput = $("xmlFile");
const xmlText = $("xmlText");
const fileStatus = $("fileStatus");
const message = $("message");
const convertButton = $("convertButton");
const progress = $("progress");
const dropzone = $("dropzone");

// ======================================================
// CONFIGURACIÓN DE NOMBRES DE HOJAS
// ======================================================

// Las secciones conocidas conservan sus nombres habituales.
// Las secciones nuevas utilizarán el nombre de su etiqueta XML.

const SHEET_NAMES = {
    compras: "Compras",
    ventas: "Ventas",
    ventasEstablecimiento: "Ventas establecimiento"
};

// ======================================================
// ETIQUETAS DE DATOS GENERALES
// ======================================================

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


// ======================================================
// ANCHOS DE LAS COLUMNAS
// ======================================================

const WIDTHS = {
    "Datos generales": [32, 40],

    "Compras": [
        13, 12, 15, 17, 12, 15, 17, 14, 12, 14,
        45, 14, 15, 13, 12, 12, 12, 14, 14, 16,
        14, 19, 15, 18, 25, 27, 31, 33, 25, 24,
        26, 27, 30, 26, 17, 18, 15, 15, 14
    ],

    "Ventas": [
        13, 15, 14, 17, 13, 20, 14, 15, 13, 12,
        12, 13, 15, 24
    ],

    "Ventas establecimiento": [12, 13, 12]
};


// ======================================================
// COLUMNAS QUE SERÁN TEXTO
// ======================================================

const TEXT_COLUMNS = {
    compras: [
        "A", "B", "C", "D", "E", "G", "H", "I",
        "K", "Y", "Z", "AA", "AB", "AC", "AI",
        "AJ", "AK", "AL"
    ],

    ventas: [
        "A", "B", "C", "D", "E"
    ],

    ventasEstablecimiento: [
        "A"
    ]
};


// ======================================================
// COLUMNAS QUE SERÁN NÚMEROS
// ======================================================

const ACCOUNTING_COLUMNS = {
    compras: [
        "L", "M", "N", "O", "P", "Q", "R", "S",
        "T", "U", "V", "W", "X", "AF", "AG", "AH"
    ],

    ventas: [
        "G", "H", "I", "J", "K", "L", "M"
    ],

    ventasEstablecimiento: [
        "B", "C"
    ]
};


// ======================================================
// COLUMNAS QUE SERÁN FECHAS
// ======================================================

const DATE_COLUMNS = {
    compras: ["J", "F", "AM"],
    ventas: [],
    ventasEstablecimiento: []
};


// ======================================================
// COLUMNAS QUE SOLAMENTE SE CENTRARÁN
// ======================================================

const CENTER_COLUMNS = {
    compras: ["J", "AD", "AE"],
    ventas: ["F", "N"],
    ventasEstablecimiento: ["A"]
};


// ======================================================
// FORMATO DE COLUMNAS
// ======================================================

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


    // FORMATO CONTABILIDAD Y DERECHA

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

                const number = Number(cell.value);

                if (Number.isFinite(number)) {
                    cell.value = number;
                }

            }

        }

    });


    // SOLO CENTRADO

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

            if (cell.value === null || cell.value === "") {
                continue;
            }

            const parts = String(cell.value).split("/");

            if (parts.length === 3) {

                const day = Number(parts[0]);
                const month = Number(parts[1]);
                const year = Number(parts[2]);

                const date = new Date(
                    Date.UTC(year, month - 1, day)
                );

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


    // LAS SECCIONES NUEVAS SE CONSERVAN COMO TEXTO
    // Y SUS COLUMNAS SE CENTRAN AUTOMÁTICAMENTE.

    if (
        !TEXT_COLUMNS[sectionKey] &&
        !ACCOUNTING_COLUMNS[sectionKey] &&
        !DATE_COLUMNS[sectionKey] &&
        !CENTER_COLUMNS[sectionKey]
    ) {

        for (let col = 1; col <= ws.columnCount; col++) {

            const column = ws.getColumn(col);

            column.numFmt = "@";

            column.alignment = {
                horizontal: "center",
                vertical: "middle"
            };

        }

    }

}


// ======================================================
// FUNCIONES PARA LEER EL XML
// ======================================================

function localName(node) {

    return node.localName || node.nodeName.split(":").pop();

}


function parseXml(xmlString) {

    const parser = new DOMParser();

    const doc = parser.parseFromString(
        xmlString,
        "application/xml"
    );

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


// ======================================================
// DETECTAR AUTOMÁTICAMENTE LAS SECCIONES
// ======================================================

function detectSections(root) {

    const sections = [];

    for (const child of Array.from(root.children)) {

        const records = Array.from(child.children);

        // Una sección de datos contiene registros que,
        // a su vez, tienen elementos internos.

        const hasRecords = records.some(record => {

            return record.children && record.children.length > 0;

        });

        if (hasRecords) {

            sections.push({
                key: localName(child),
                node: child,
                records: records.filter(record => {
                    return record.children && record.children.length > 0;
                })
            });

        }

    }

    return sections;

}


// ======================================================
// CONVERTIR XML ANIDADO EN COLUMNAS
// ======================================================

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


// ======================================================
// CREAR ENCABEZADOS Y FILAS
// ======================================================

function makeRows(records) {

    const flattened = records.map(record => {
        return flattenElement(record);
    });

    const headers = [];

    for (const record of flattened) {

        for (const key of Object.keys(record)) {

            if (!headers.includes(key)) {
                headers.push(key);
            }

        }

    }

    const rows = flattened.map(record => {

        return headers.map(header => {
            return record[header] ?? "";
        });

    });

    return {
        headers,
        rows
    };

}


// ======================================================
// ESTILO DE ENCABEZADOS
// ======================================================

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


// ======================================================
// CONFIGURAR DISEÑO DE LAS HOJAS
// ======================================================

function setSheetLayout(ws, sheetName, headers) {

    const widths = WIDTHS[sheetName] || [];

    ws.columns = headers.map((header, index) => ({

        header,

        key: `c${index}`,

        width: widths[index] ||
            Math.min(
                Math.max(String(header).length + 3, 14),
                32
            )

    }));

    setHeader(ws.getRow(1));

    ws.views = [
        {
            state: "frozen",
            ySplit: 1
        }
    ];

    if (headers.length && ws.rowCount > 1) {

        ws.autoFilter = {
            from: {
                row: 1,
                column: 1
            },
            to: {
                row: ws.rowCount,
                column: headers.length
            }
        };

    }

}


// ======================================================
// EVITAR NOMBRES DE HOJAS NO VÁLIDOS O REPETIDOS
// ======================================================

function getUniqueSheetName(workbook, name) {

    let safeName = String(name)
        .replace(/[\\/*?:[\]]/g, "_")
        .substring(0, 31);

    if (!safeName) {
        safeName = "Sección";
    }

    let finalName = safeName;
    let counter = 2;

    while (
        workbook.worksheets.some(ws => {
            return ws.name.toLowerCase() === finalName.toLowerCase();
        })
    ) {

        const suffix = ` (${counter})`;

        finalName =
            safeName.substring(0, 31 - suffix.length) + suffix;

        counter++;

    }

    return finalName;

}


// ======================================================
// CREAR HOJA DE DATOS GENERALES
// ======================================================

function addGeneralSheet(workbook, root, sections) {

    const ws = workbook.addWorksheet("Datos generales");

    ws.addRow(["Campo", "Valor"]);

    // Identifica las secciones que contienen registros.

    const sectionNodes = new Set(
        sections.map(section => section.node)
    );

    for (const child of Array.from(root.children)) {

        // No incluir las secciones de datos aquí.

        if (sectionNodes.has(child)) {
            continue;
        }

        const key = localName(child);

        ws.addRow([
            GENERAL_LABELS[key] || key,
            (child.textContent || "").trim()
        ]);

    }

    setSheetLayout(
        ws,
        "Datos generales",
        ["Campo", "Valor"]
    );

    ws.getColumn(1).width = 32;
    ws.getColumn(2).width = 40;

    for (let row = 2; row <= ws.rowCount; row++) {

        ws.getCell(row, 2).numFmt = "@";

        ws.getCell(row, 2).alignment = {
            horizontal: "center",
            vertical: "middle"
        };

    }

}


// ======================================================
// CREAR UNA HOJA POR CADA SECCIÓN DEL XML
// ======================================================

function addDataSheet(workbook, section) {

    const sectionKey = section.key;

    const sheetTitle =
        SHEET_NAMES[sectionKey] || sectionKey;

    const uniqueName = getUniqueSheetName(
        workbook,
        sheetTitle
    );

    const { headers, rows } = makeRows(section.records);

    const ws = workbook.addWorksheet(uniqueName);

    if (!headers.length) {

        ws.addRow([
            "No se encontraron registros en esta sección"
        ]);

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

    setSheetLayout(
        ws,
        uniqueName,
        headers
    );

    // Conserva inicialmente todos los valores como texto.

    for (let row = 2; row <= ws.rowCount; row++) {

        for (let col = 1; col <= headers.length; col++) {

            const cell = ws.getRow(row).getCell(col);

            if (
                cell.value !== null &&
                cell.value !== undefined
            ) {

                cell.value = String(cell.value);

            }

        }

    }

    // Aplica los formatos correspondientes.

    applyColumnFormats(ws, sectionKey);

}


// ======================================================
// CONSTRUIR EL ARCHIVO EXCEL
// ======================================================

function buildWorkbook(xmlString) {

    const doc = parseXml(xmlString);

    const root = doc.documentElement;

    const workbook = new ExcelJS.Workbook();

    workbook.creator = "Conversor XML ATS";

    workbook.subject = "Conversión de XML del ATS a Excel";

    workbook.created = new Date();

    // Detecta todas las secciones del XML.

    const sections = detectSections(root);

    // Primera hoja: datos generales.

    addGeneralSheet(
        workbook,
        root,
        sections
    );

    // Las demás hojas se crean automáticamente.

    for (const section of sections) {

        addDataSheet(
            workbook,
            section
        );

    }

    return {
        workbook,
        sections
    };

}


// ======================================================
// NOMBRE DEL ARCHIVO EXCEL
// ======================================================

function makeOutputName(fileName = "") {

    if (fileName) {

        return fileName.replace(/\.xml$/i, "") + ".xlsx";

    }

    return "ReporteATS_convertido.xlsx";

}


// ======================================================
// DESCARGAR EL EXCEL
// ======================================================

async function downloadWorkbook(workbook, fileName) {

    const buffer = await workbook.xlsx.writeBuffer();

    const blob = new Blob(
        [buffer],
        {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
    );

    const url = URL.createObjectURL(blob);

    const link = document.createElement("a");

    link.href = url;

    link.download = makeOutputName(fileName);

    document.body.appendChild(link);

    link.click();

    link.remove();

    setTimeout(() => {
        URL.revokeObjectURL(url);
    }, 1500);

}


// ======================================================
// PROCESAR UN XML
// ======================================================

async function processXmlContent(content, fileName = "") {

    const result = buildWorkbook(content);

    await downloadWorkbook(
        result.workbook,
        fileName
    );

    const counts = result.sections.map(section => {

        return `${SHEET_NAMES[section.key] || section.key}: ${section.records.length}`;

    }).join(" · ");

    return counts || "No se encontraron secciones de datos adicionales.";

}


// ======================================================
// CONVERTIR Y DESCARGAR
// ======================================================

async function convertAndDownload() {

    hideMessage();

    const files = Array.from(
        fileInput.files || []
    );

    const pastedContent = xmlText.value.trim();

    // Si no hay archivos seleccionados, utiliza el XML pegado.

    if (!files.length && !pastedContent) {

        showMessage(
            "Selecciona un archivo XML o pega primero su contenido.",
            "error"
        );

        return;

    }

    convertButton.disabled = true;

    progress.hidden = false;

    try {

        const results = [];

        // PROCESAR LOS ARCHIVOS SELECCIONADOS.

        if (files.length) {

            for (let i = 0; i < files.length; i++) {

                const file = files[i];

                if (!/\.xml$/i.test(file.name) && !/xml/i.test(file.type)) {

                    throw new Error(
                        `El archivo "${file.name}" no tiene formato XML.`
                    );

                }

                const content = await file.text();

                const counts = await processXmlContent(
                    content,
                    file.name
                );

                results.push(
                    `${file.name}: ${counts}`
                );

                // Pequeña pausa entre descargas múltiples.

                await new Promise(resolve => {
                    setTimeout(resolve, 300);
                });

            }

        } else {

            // PROCESAR EL XML PEGADO MANUALMENTE.

            const counts = await processXmlContent(
                pastedContent
            );

            results.push(counts);

        }

        showMessage(
            `¡Excel generado correctamente! ${results.join(" | ")}`,
            "success"
        );

    } catch (error) {

        console.error(error);

        showMessage(
            error.message || "No se pudo generar el archivo Excel.",
            "error"
        );

    } finally {

        convertButton.disabled = false;

        progress.hidden = true;

    }

}


// ======================================================
// MENSAJES
// ======================================================

function showMessage(text, type) {

    message.textContent = text;

    message.className = `message ${type}`;

}


function hideMessage() {

    message.textContent = "";

    message.className = "message";

}


// ======================================================
// SELECCIÓN DE ARCHIVOS
// ======================================================

fileInput.addEventListener("change", async () => {

    const files = Array.from(
        fileInput.files || []
    );

    if (!files.length) {
        return;
    }

    try {

        // Muestra el contenido del primer XML en el cuadro de texto.

        xmlText.value = await files[0].text();

        if (files.length === 1) {

            fileStatus.textContent =
                `Archivo cargado: ${files[0].name} (${Math.ceil(files[0].size / 1024)} KB)`;

        } else {

            fileStatus.textContent =
                `${files.length} archivos XML seleccionados. Se generará un Excel por cada archivo.`;

        }

        hideMessage();

    } catch {

        showMessage(
            "No se pudo leer el archivo. Intenta abrirlo y pegar su contenido.",
            "error"
        );

    }

});


// ======================================================
// BOTÓN LIMPIAR
// ======================================================

$("clearButton").addEventListener("click", () => {

    fileInput.value = "";

    xmlText.value = "";

    fileStatus.textContent =
        "Todavía no has cargado ningún XML.";

    hideMessage();

});


// ======================================================
// BOTÓN CONVERTIR
// ======================================================

convertButton.addEventListener(
    "click",
    convertAndDownload
);


// ======================================================
// ARRASTRAR Y SOLTAR
// ======================================================

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

    const files = Array.from(
        event.dataTransfer.files || []
    );

    if (!files.length) {
        return;
    }

    // Comprueba que todos los archivos sean XML.

    const invalidFile = files.find(file => {

        return !/\.xml$/i.test(file.name) &&
            !/xml/i.test(file.type);

    });

    if (invalidFile) {

        showMessage(
            `El archivo "${invalidFile.name}" no tiene formato XML.`,
            "error"
        );

        return;

    }

    try {

        // Permite seleccionar varios archivos arrastrados.

        const dataTransfer = new DataTransfer();

        files.forEach(file => {
            dataTransfer.items.add(file);
        });

        fileInput.files = dataTransfer.files;

        // Muestra el contenido del primer XML.

        xmlText.value = await files[0].text();

        if (files.length === 1) {

            fileStatus.textContent =
                `Archivo cargado: ${files[0].name} (${Math.ceil(files[0].size / 1024)} KB)`;

        } else {

            fileStatus.textContent =
                `${files.length} archivos XML seleccionados. Se generará un Excel por cada archivo.`;

        }

        hideMessage();

    } catch (error) {

        console.error(error);

        showMessage(
            "No se pudieron cargar los archivos.",
            "error"
        );

    }

});