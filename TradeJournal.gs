const tradeJournalFirstRow = 4;
const exportConfigSheetName = "Export";
const instrumentsConfigSheetName = "Инструменты";

const tradeJournalSheetConfiguation = [
  {
    csvTitle: "Date",
    propertyName: "date",
    converter: dateConverter,
    columnName: "A",
  },
  {
    csvTitle: "Time",
    propertyName: "time",
    valueGetter: (rowRange, cell, sheetConfig, instrumentsConfig) => cell.getDisplayValue(),
    columnName: "B",
  },
  {
    csvTitle: "Type",
    propertyName: "type",
    valueGetter: typeValueGetter_,
    columnName: "H",
  },
  {
    csvTitle: "Value",
    propertyName: "value",
    converter: moneyConverter,
    valueGetter: valueGetter_,
  },
  {
    csvTitle: "Security Name",
    propertyName: "name",
    valueGetter: securityNameValueGetter_,
    columnName: "C",
  },
  {
    csvTitle: "Shares",
    propertyName: "shares",
    columnName: "K",
  },
  {
    csvTitle: "Quote",
    propertyName: "quote",
    converter: moneyConverter,
    columnName: "I",
  },
  {
    csvTitle: "Offset Account",
    propertyName: "offsetAccount",
    valueGetter: offsetAccountValueGetter_,
  },
  {
    csvTitle: "Note",
    propertyName: "note",
    valueGetter: noteValueGetter_,
    columnName: "C",
  },
  {
    csvTitle: "Fees",
    propertyName: "fees",
    converter: moneyConverter,
    columnName: "R",
  },
  {
    csvTitle: "Taxes",
    propertyName: "taxes",
    converter: moneyConverter,
    valueGetter: (rowRange, cell, sheetConfig, instrumentsConfig) => null,
  },
  {
    csvTitle: "Cash Account",
    propertyName: "cashAccount",
    valueGetter: cashAccountValueGetter_,
  },
  {
    csvTitle: "Securities Account",
    propertyName: "securitiesAccount",
    valueGetter: securitiesAccountValueGetter_,
  },
  {
    csvTitle: "Transaction Currency",
    propertyName: "transactionCurrency",
    valueGetter: (rowRange, cell, sheetConfig, instrumentsConfig) => "RUB",
  },
]

function getCsvFromTradeSheet(sheet, startRow, endRow, config, delimiter, state) {
  if (startRow < tradeJournalFirstRow || startRow > sheet.getMaxRows()) {
    throw new RangeError("Start row value is out of range");
  }
  if (endRow < tradeJournalFirstRow || endRow > sheet.getMaxRows()) {
    throw new RangeError("End row value is out of range");
  }
  if (startRow > endRow) {
    throw new RangeError("Start row is greater then end row");
  }

  const sheetName = sheet.getName();
  const sheetConfig = config.sheets.find((item) => item.sheetName === sheetName);
  if (sheetConfig === undefined) {
    throw new Error(`Sheet ${sheetName} is not configured for export`);
  }

  const stepsCount = endRow - startRow + 1;
  const cache = CacheService.getScriptCache();

  state.msg = "Generate header";
  state.prc = Math.trunc(1.0 * 100.0 / stepsCount);
  propagateState(state, cache);

  const sheetMaxColumns = sheet.getMaxColumns();
  const instrumentsConfig = config.instruments;

  // prepare CSV header data
  var headerData = Array.from(tradeJournalSheetConfiguation, (config) => {
    var headerItem = {
      title: config.csvTitle,
      propertyName: config.propertyName,
    };

    if (config.converter) {
      headerItem.converter = config.converter;
    }

    return headerItem;
  });

  // prepare data to export
  var dataToExport = [];
  for (var row = startRow; row <= endRow; row++) {
    state.msg = `Export row ${row}`;
    state.prc = Math.trunc((row + 1) * 100.0 / stepsCount);
    propagateState(state, cache);

    // entire row
    const rowRange = sheet.getRange(row, 1, 1, sheetMaxColumns);

    // item to build
    var itemToExport = {};

    tradeJournalSheetConfiguation.forEach((configItem) => {
      const cell = configItem.columnName ? getRangeByColName(sheet, row, configItem.columnName) : rowRange;
      const cellValue = configItem.valueGetter ? configItem.valueGetter(rowRange, cell, sheetConfig, instrumentsConfig) : cell.getValue();
      
      itemToExport[configItem.propertyName] = cellValue;
    });

    dataToExport.push(itemToExport);
  }

  state.msg = "Generate CSV text";
  state.prc = 100;
  propagateState(state, cache);

  const csv = getCsvFromJson(headerData, dataToExport, delimiter);
  return csv;
}

function getTradeSheetsConfig(exportConfigSheet, instrumentsConfigSheet) {
  // read export sheets config
  var exportConfigData = exportConfigSheet.getDataRange().getValues();

  var sheetsConfig = [];
  var configProperties = ["sheetName", "securitiesAccount", "depositAccount"];

  // first row is header
  for (var row = 1; row < exportConfigData.length; row++) {
    var sheetData = {};
    for (var col = 0; col < configProperties.length; col++) {
      sheetData[configProperties[col]] = exportConfigData[row][col];
    }
    sheetsConfig.push(sheetData);
  }

  // read instruments config
  var instrumentsConfigData = instrumentsConfigSheet.getDataRange().getValues();

  var instrumentsConfig = [];
  configProperties = ["name", "isin", "ticker", "type", "currency", "exportName"];

  // first row is header
  for (var row = 1; row < instrumentsConfigData.length; row++) {
    var instrumentData = {};
    for (var col = 0; col < configProperties.length; col++) {
      instrumentData[configProperties[col]] = instrumentsConfigData[row][col];
    }
    instrumentsConfig.push(instrumentData);
  }

  var config = {
    sheets: sheetsConfig,
    instruments: instrumentsConfig,
  };

  return config;
}

function getDatesFromSheet(sheet) {
  const lastRow = sheet.getDataRange().getLastRow();
  var datesInSheet = {};
  for (var row = tradeJournalFirstRow; row <= lastRow; row++) {
    // entire row
    const rowRange = sheet.getRange(row, 1, 1, sheet.getMaxColumns());
    
    const dateValue = getCellValueByPropertyName_(rowRange, "date");
    if (dateValue) {
      const dateKey = dateValue.toISOString();

      if (dateKey in datesInSheet) {
        var existingItem = datesInSheet[dateKey];
        existingItem.lastRowNum = row;
      }
      else {
        datesInSheet[dateKey] = {date: dateValue, firstRowNum: row, lastRowNum: row};      
      }
    }
  }

  datesInSheet = Object.values(datesInSheet);
  datesInSheet.sort((it1, it2) => it1.firstRowNum - it2.firstRowNum);
  return datesInSheet;
}

function getCellValueByPropertyName_(rowRange, propertyName) {
  const configItem = tradeJournalSheetConfiguation.find((configItem) => configItem.propertyName === propertyName);

  if (configItem.columnName) {
    const sheet = rowRange.getSheet();
    const row = rowRange.getRow();
    const cell = getRangeByColName(sheet, row, configItem.columnName);
    return cell.getValue();
  }

  return undefined;
}

function typeValueGetter_(rowRange, cell, sheetConfig, instrumentsConfig) {
  const typeValue = cell.getValue();

  switch(typeValue) {
    case "Buy":
      return "Buy";
    case "Sell":
      return "Sell";
    case "Внесение":
      return "Deposit";
    case "Вывод":
      return "Removal";
    case "Купоны":
      return "Dividend";
    case "Дивиденды":
      return "Dividend";
    case "НДФЛ":
      return "Taxes";
    case "Возврат НДФЛ":
      return "Tax Refund";
    case "Депо. комиссия":
      return "Fees";
    case "Комиссия":
      return "Fees";
    case "Проценты":
      return "Interest";
    case "Возврат процентов":
      return "Interest Charge";
    default:
      throw new Error(`Unknown operation type ${typeValue}`);
  }
}

function securityNameValueGetter_(rowRange, cell, sheetConfig, instrumentsConfig) {
  const cellValue = cell.getValue();
  const typeValue = getCellValueByPropertyName_(rowRange, "type");

  switch(typeValue) {
    case "Buy":
    case "Sell":
    case "Купоны":
    case "Дивиденды":
      if (!cellValue) {
        return null;
      }

      const instrumentConfig = instrumentsConfig.find((instrItem) => instrItem.name === cellValue);
      if (instrumentConfig === undefined) {
        throw new Error(`Instrument '${cellValue}' not found`);
      }
      
      return instrumentConfig.exportName;

    case "Внесение":
    case "Вывод":
    case "НДФЛ":
    case "Возврат НДФЛ":
    case "Депо. комиссия":
    case "Комиссия":
    case "Проценты":
    case "Возврат процентов":
      return null;

    default:
      throw new Error(`Unknown operation type ${typeValue}`);
  }
}

function noteValueGetter_(rowRange, cell, sheetConfig, instrumentsConfig) {
  const cellValue = cell.getValue();
  const typeValue = getCellValueByPropertyName_(rowRange, "type");

  switch(typeValue) {
    case "Buy":
    case "Sell":
    case "Купоны":
    case "Дивиденды":
      return null;

    case "Внесение":
    case "Вывод":
    case "НДФЛ":
    case "Возврат НДФЛ":
    case "Депо. комиссия":
    case "Комиссия":
    case "Проценты":
    case "Возврат процентов":
      return cellValue ? cellValue : null;

    default:
      throw new Error(`Unknown operation type ${typeValue}`);
  }
}

function valueGetter_(rowRange, cell, sheetConfig, instrumentsConfig) {
  const sheet = rowRange.getSheet();
  const row = rowRange.getRow();
  const typeValue = getCellValueByPropertyName_(rowRange, "type");

  var cell;
  switch(typeValue) {
    case "Buy":
    case "Sell":
      cell = getRangeByColName(sheet, row, "T");
      break;

    case "Внесение":
    case "Вывод":
    case "Купоны":
    case "Дивиденды":
    case "НДФЛ":
    case "Возврат НДФЛ":
    case "Депо. комиссия":
    case "Комиссия":
    case "Проценты":
    case "Возврат процентов":
      cell = getRangeByColName(sheet, row, "U");
      break;

    default:
      throw new Error(`Unknown operation type ${typeValue}`);
  }

  return cell.getValue();
}

function offsetAccountValueGetter_(rowRange, cell, sheetConfig, instrumentsConfig) {
  return sheetConfig.securitiesAccount;
}

function cashAccountValueGetter_(rowRange, cell, sheetConfig, instrumentsConfig) {
  return sheetConfig.depositAccount;
}

function securitiesAccountValueGetter_(rowRange, cell, sheetConfig, instrumentsConfig) {
  return sheetConfig.securitiesAccount;
}

function getRangeByColName(sheet, row, columnName) {
  return sheet.getRange(`${columnName}${row}`);
}
