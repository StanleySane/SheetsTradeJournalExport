function getCsvFromJson(headerData, dataToExport, delimiter) {
  var headerItems = Array.from(headerData, (headerItem) => headerItem.title);
  var csv = headerItems.join(delimiter) + "\r\n";

  if (dataToExport.length == 0) {
    return csv;
  }
  
  for (var itemIndex = 0; itemIndex < dataToExport.length; itemIndex++) {
    var itemData = dataToExport[itemIndex];
    var rowData = Array.from(headerData, (headerItem) => {
      var propertyValue = itemData[headerItem.propertyName];
      var propertyStringValue = headerItem.converter
        ? headerItem.converter(propertyValue)
        : (propertyValue === null ? "" : propertyValue.toString());

      return escape(propertyStringValue, delimiter);
    });

    csv += rowData.join(delimiter);

    if (itemIndex < dataToExport.length - 1) {
      csv += "\r\n";
    }
  }

  return csv;
}

function dateConverter(dateValue) {
  // drop timezone
  const localeDate = new Date(Date.UTC(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate()));
  return localeDate.toISOString().slice(0, 10);
}

function moneyConverter(moneyValue) {
  return moneyValue === null ? "" : moneyValue.toString();
}

function escape(stringValue, delimiter) {
  if (stringValue.indexOf(delimiter) != -1) {
    return `"${stringValue.replace(/"/g, '""')}"`;
  }
  return stringValue;
}

function convertRangeToCsvFile(sheet, delimiter) {
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  try {
    var data = activeRange.getValues();
    var csvFile = undefined;

    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(delimiter) != -1) {
            data[row][col] = (`"${data[row][col].replace(/"/g, '""')}"`) ;
          }
        }

        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(delimiter) + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
    }

    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}

