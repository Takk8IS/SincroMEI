//```Code.gs

/**
 * Copyright (c) Takk™ Innovate Studio
 * SincroMEI Google Sheets Add-on
 * Description: This Google Sheets add-on allows you to synchronise data from the SincroMEI API.
 * Author: David C Cavalcante
 * Version: 1.0.0
 * License: CC-BY-4.0
 * Repository: https://github.com/takk8is/sincromei.git
 * Donate: $USDT (TRC-20) TGpiWetnYK2VQpxNGPR27D9vfM6Mei5vNA
 */

/**
 * Constants and configuration
 */
const SHEET_NAME = "SincroMEI Data";
const API_BASE_URL = "http://0.0.0.0:8002";
const HEADERS = [
    "CNPJ",
    "Situação",
    "Última Atualização",
    "Razão Social",
    "Tipo",
    "Porte",
    "Abertura",
    "Natureza Jurídica",
    "Logradouro",
    "Número",
    "CEP",
    "Bairro",
    "Município",
    "UF",
    "Email",
    "Telefone",
    "Capital Social",
    "2019",
    "2020",
    "2021",
    "2022",
    "2023",
    "2024",
];

let isRunning = false;
let START_TIME;

/**
 * Creates a menu item for the add-on when the spreadsheet opens.
 */
function onOpen(e) {
    if (e && e.authMode !== ScriptApp.AuthMode.NONE) {
        SpreadsheetApp.getUi()
            .createMenu("SincroMEI")
            .addItem("Configurar Planilha", "configureSpreadsheet")
            .addItem("Iniciar Sincronização", "startSynchronisation")
            .addItem("Pausar Sincronização", "pauseSynchronisation")
            .addToUi();
    }
}

/**
 * Configures the spreadsheet with initial settings.
 */
function configureSpreadsheet() {
    const sheet = getOrCreateSheet(SHEET_NAME);
    formatSheet(sheet);
    configureHeaders(sheet);
    formatCurrencyColumn(sheet);
    SpreadsheetApp.getUi().alert("A planilha foi configurada com sucesso!");
}

/**
 * Starts the synchronisation process.
 */
function startSynchronisation() {
    if (!isRunning) {
        isRunning = true;
        START_TIME = new Date().getTime();
        synchroniseData();
    } else {
        SpreadsheetApp.getUi().alert("A sincronização já está em andamento.");
    }
}

/**
 * Pauses the synchronisation process.
 */
function pauseSynchronisation() {
    if (isRunning) {
        isRunning = false;
        SpreadsheetApp.getUi().alert(
            "A sincronização foi pausada. Você pode retomá-la selecionando 'Iniciar Sincronização' novamente.",
        );
    } else {
        SpreadsheetApp.getUi().alert(
            "Não há sincronização em andamento para pausar.",
        );
    }
}

/**
 * Main function to synchronise data from the API.
 */
function synchroniseData() {
    const sheet = getOrCreateSheet(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const lastRow = sheet.getLastRow();

    const progressCell = sheet.getRange("A1");
    progressCell.setValue("Iniciando sincronização...");

    // Process 3 CNPJs per minute
    const batchSize = 3;
    // 1 minute delay between batches
    const delay = 60000;

    for (let i = 1; i < data.length && isRunning; i += batchSize) {
        const batch = data.slice(i, i + batchSize);
        processBatch(sheet, batch, i);

        if (i + batchSize < data.length) {
            progressCell.setValue(
                `Processados ${i + batchSize} de ${data.length} CNPJs. Aguardando próximo lote...`,
            );
            Utilities.sleep(delay);
        }

        // Check if we're approaching the maximum execution time
        if (new Date().getTime() - START_TIME > 350000) {
            // 350 seconds (just under 6 minutes)
            progressCell.setValue(
                `Tempo máximo de execução atingido. Retome a sincronização para continuar.`,
            );
            isRunning = false;
            break;
        }
    }

    if (isRunning) {
        progressCell.setValue("Sincronização concluída!");
        SpreadsheetApp.getUi().alert("Sincronização concluída!");
    }

    isRunning = false;
}

/**
 * Processes a batch of CNPJs.
 */
function processBatch(sheet, batch, startIndex) {
    batch.forEach((row, index) => {
        const cnpj = row[0];
        if (cnpj && !isProcessed(sheet, startIndex + index + 1)) {
            const result = processCNPJ(cnpj);
            if (result) {
                writeDataToSheet(sheet, startIndex + index, result);
            } else {
                markAsError(sheet, startIndex + index + 1);
            }
        }
    });
}

/**
 * Checks if a row has already been processed.
 */
function isProcessed(sheet, row) {
    return sheet
        .getRange(row, 2, 1, HEADERS.length - 1)
        .getValues()[0]
        .some((cell) => cell !== "");
}

/**
 * Marks a row as having an error in fetching data.
 */
function markAsError(sheet, row) {
    sheet
        .getRange(row, 2, 1, HEADERS.length - 1)
        .setValue("Erro ao buscar dados");
}

/**
 * Configures the header row of the sheet.
 */
function configureHeaders(sheet) {
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);
    headerRange
        .setBackground("#158821")
        .setFontColor("white")
        .setFontWeight("bold")
        .setFontFamily("Inter")
        .setFontSize(10)
        .setValues([HEADERS.map((header) => header.toUpperCase())])
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
}

/**
 * Formats the entire sheet.
 */
function formatSheet(sheet) {
    // Ensure at least two rows
    const lastRow = Math.max(sheet.getLastRow(), 2);
    const lastColumn = HEADERS.length;
    const range = sheet.getRange(1, 1, lastRow, lastColumn);

    // Set column width and row height
    sheet.setColumnWidths(1, lastColumn, 300);
    sheet.setRowHeights(1, lastRow, 70);

    // Set text wrapping and alignment
    range.setWrap(true);
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");

    // Remove borders
    range.setBorder(false, false, false, false, false, false);

    // Set font for all cells
    range.setFontFamily("Inter");
    range.setFontSize(10);

    // Set alternating row colors starting from row 2
    for (let i = 2; i <= lastRow; i++) {
        const rowRange = sheet.getRange(i, 1, 1, lastColumn);
        if (i % 2 === 0) {
            rowRange.setBackground("white");
        } else {
            // Light gray
            rowRange.setBackground("#F8F9FA");
        }
        rowRange.setFontColor("black");
    }
}

/**
 * Formats CNPJ by removing non-digit characters.
 */
function formatCNPJ(cnpj) {
    return cnpj.toString().replace(/[^\d]/g, "");
}

/**
 * Fetches data for a single CNPJ from the API.
 */
function fetchRevenueData(cnpj) {
    const url = `${API_BASE_URL}/sincromei/${cnpj}`;
    try {
        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const responseCode = response.getResponseCode();

        if (responseCode === 200) {
            return JSON.parse(response.getContentText());
        } else {
            Logger.log(`API Error: ${responseCode}`);
            return null;
        }
    } catch (e) {
        Logger.log(`Error fetching data: ${e.message}`);
        return null;
    }
}

/**
 * Processes years data from the API response.
 */
function processYears(anos) {
    const currentYear = new Date().getFullYear();
    const processedYears = {};

    for (let year = 2019; year <= currentYear; year++) {
        const yearData = anos[year];
        processedYears[year] = yearData ? yearData : "Não Optante";
    }

    return processedYears;
}

/**
 * Processes a single CNPJ, fetching and formatting its data.
 */
function processCNPJ(cnpj) {
    const formattedCNPJ = formatCNPJ(cnpj);
    const result = fetchRevenueData(formattedCNPJ);
    if (result) {
        const processedYears = processYears(result.declaracao || {});
        return {
            cnpj: formattedCNPJ,
            situacao: result.situacao,
            ultimaAtualizacao: result.ultima_atualizacao,
            razaoSocial: result.nome,
            tipo: result.tipo,
            porte: result.porte,
            abertura: result.abertura,
            naturezaJuridica: result.natureza_juridica,
            logradouro: result.logradouro,
            numero: result.numero,
            cep: result.cep,
            bairro: result.bairro,
            municipio: result.municipio,
            uf: result.uf,
            email: result.email,
            telefone: result.telefone,
            capitalSocial: result.capital_social,
            anos: processedYears,
        };
    }
    return null;
}

/**
 * Writes processed data to the sheet.
 */
function writeDataToSheet(sheet, rowIndex, data) {
    const row = rowIndex + 1;
    sheet.getRange(row, 1).setValue(data.cnpj);
    sheet.getRange(row, 2).setValue(data.situacao);
    sheet.getRange(row, 3).setValue(data.ultimaAtualizacao);
    sheet.getRange(row, 4).setValue(data.razaoSocial);
    sheet.getRange(row, 5).setValue(data.tipo);
    sheet.getRange(row, 6).setValue(data.porte);
    sheet.getRange(row, 7).setValue(data.abertura);
    sheet.getRange(row, 8).setValue(data.naturezaJuridica);
    sheet.getRange(row, 9).setValue(data.logradouro);
    sheet.getRange(row, 10).setValue(data.numero);
    sheet.getRange(row, 11).setValue(data.cep);
    sheet.getRange(row, 12).setValue(data.bairro);
    sheet.getRange(row, 13).setValue(data.municipio);
    sheet.getRange(row, 14).setValue(data.uf);
    sheet.getRange(row, 15).setValue(data.email);
    sheet.getRange(row, 16).setValue(data.telefone);
    sheet.getRange(row, 17).setValue(data.capitalSocial);

    formatCurrencyCell(sheet, row, 17);

    for (let year = 2019; year <= 2024; year++) {
        sheet
            .getRange(row, year - 2001)
            .setValue(data.anos[year] || "Não Optante");
    }
}

/**
 * Formats the entire "Capital Social" column as currency in R$.
 */
function formatCurrencyColumn(sheet) {
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(2, 17, lastRow - 1, 1);
    range.setNumberFormat("R$ #,##0.00");
}

/**
 * Formats a specific cell in the "Capital Social" column as currency in R$.
 */
function formatCurrencyCell(sheet, row, column) {
    const range = sheet.getRange(row, column);
    range.setNumberFormat("R$ #,##0.00");
}

/**
 * Gets an existing sheet or creates a new one if it doesn't exist.
 */
function getOrCreateSheet(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
    }
    return sheet;
}

/**
 * Sets environment variables for the script.
 */
function setEnvironmentVariables() {
    PropertiesService.getScriptProperties().setProperties({
        API_URL: API_BASE_URL.split(":")[0] + ":" + API_BASE_URL.split(":")[1],
        API_PORT: API_BASE_URL.split(":")[2],
    });
}

// Run this function once to set the environment variables
setEnvironmentVariables();
