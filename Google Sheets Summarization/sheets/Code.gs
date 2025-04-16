URL = "<YOUR_NGROK_URL_HERE>"

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('AI Summarization')
      .addItem('Open Sidebar', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('AI Summarization');
  SpreadsheetApp.getUi().showSidebar(html);
}

const ALLOWED_STYLES = ["academic", "creative", "formal/business", "informal/casual"];
const ALLOWED_MODELS = ["gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-1.5-flash", "gemini-1.5-pro"];

function processRows(settings) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  if (settings.processMode === 'range') {
    // Range Mode: Process specific cell ranges
    var sourceRange = parseRange(settings.sourceRange, lastRow, lastColumn);
    var resultRange = parseRange(settings.resultRange, lastRow, lastColumn);

    // Validate range compatibility
    if (sourceRange.cells.length !== resultRange.cells.length) {
      throw new Error('Source and result ranges must have the same number of cells');
    }

    // Get source range values
    var sourceValues = sheet.getRange(
      sourceRange.startRow, 
      sourceRange.startCol, 
      sourceRange.numRows, 
      sourceRange.numCols
    ).getValues();

    // Process each cell in the source range
    var results = [];
    for (var i = 0; i < sourceRange.cells.length; i++) {
      var cell = sourceRange.cells[i];
      var rowIndex = cell.row - sourceRange.startRow;
      var colIndex = cell.col - sourceRange.startCol;
      var text = sourceValues[rowIndex][colIndex] || '';
      
      // Skip empty cells
      if (!text) {
        results.push('');
        continue;
      }

      // Call the API
      var response = callAPI(text, settings.style, settings.model, settings.custom);
      results.push(response);
    }

    // Write results to the result range, using the result range's shape
    var resultValues = [];
    var cellIndex = 0;
    for (var i = 0; i < resultRange.numRows; i++) {
      var currentRow = [];
      for (var j = 0; j < resultRange.numCols; j++) {
        currentRow.push(cellIndex < results.length ? results[cellIndex] : '');
        cellIndex++;
      }
      resultValues.push(currentRow);
    }

    sheet.getRange(
      resultRange.startRow, 
      resultRange.startCol, 
      resultRange.numRows, 
      resultRange.numCols
    ).setValues(resultValues);

  } else {
    // Auto Columns Mode: Automatically detect non-empty cells in columns
    var sourceCols = parseColumns(settings.sourceColumns);
    var resultCols = parseColumns(settings.resultColumns);

    // Validate column compatibility
    if (sourceCols.length !== resultCols.length) {
      throw new Error('Source and result columns must have the same number of columns');
    }

    // Start processing from the row after headerRow
    var startRow = (settings.headerRow || 0) + 1;
    if (startRow > lastRow) {
      throw new Error('Header row exceeds sheet data');
    }

    // Get the full data range for source columns
    var sourceRanges = sourceCols.map(function(col) {
      return sheet.getRange(startRow, col, lastRow - startRow + 1, 1);
    });
    var sourceValues = sourceRanges.map(function(range) {
      return range.getValues().map(function(row) {
        return row[0] || '';
      });
    });

    // Find the maximum number of non-empty rows across source columns
    var maxRows = 0;
    sourceValues.forEach(function(colValues) {
      var lastNonEmpty = colValues.reduce(function(last, val, idx) {
        return val ? idx : last;
      }, -1);
      maxRows = Math.max(maxRows, lastNonEmpty + 1);
    });

    if (maxRows === 0) {
      throw new Error('No non-empty cells found in source columns');
    }

    // Process each non-empty cell
    var results = sourceCols.map(function(col, colIdx) {
      var colResults = new Array(maxRows).fill('');
      var colValues = sourceValues[colIdx];
      for (var i = 0; i < maxRows; i++) {
        var text = colValues[i];
        if (text) {
          colResults[i] = callAPI(text, settings.style, settings.model, settings.custom);
        }
      }
      return colResults;
    });

    // Write results to result columns
    resultCols.forEach(function(col, colIdx) {
      sheet.getRange(startRow, col, maxRows, 1).setValues(
        results[colIdx].map(function(val) {
          return [val];
        })
      );
    });
  }
}

function callAPI(text, style, model, custom) {
  var url = URL + '/chat'; // Matches your FastAPI server
  var payload = {
    model: model,
    text: text,
    style: style,
    custom: custom || "" // Ensure custom is an empty string if not provided
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  try {
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());
    return result.message;
  } catch (e) {
    throw new Error('API call failed: ' + e.message);
  }
}

function parseRange(rangeStr, maxRow, maxCol) {
  // Normalize range string (e.g., "A1:B10", "C:C", "A1")
  rangeStr = rangeStr.toUpperCase().trim();
  var parts = rangeStr.split(':');
  var startCell, endCell;

  if (parts.length === 1) {
    // Single cell or column (e.g., "A1" or "C")
    startCell = parts[0];
    endCell = startCell;
  } else {
    // Range (e.g., "A1:B10" or "C:D")
    startCell = parts[0];
    endCell = parts[1];
  }

  // Parse start and end cells
  var start = parseCell(startCell);
  var end = parseCell(endCell || startCell);

  // Handle column-only ranges (e.g., "C:C")
  if (!start.row) {
    start.row = 1;
    end.row = maxRow;
  }
  if (!end.row) {
    end.row = maxRow;
  }

  // Calculate dimensions
  var startRow = Math.min(start.row, end.row);
  var endRow = Math.max(start.row, end.row);
  var startCol = Math.min(start.col, end.col);
  var endCol = Math.max(start.col, end.col);

  var numRows = endRow - startRow + 1;
  var numCols = endCol - startCol + 1;

  // Generate list of cells
  var cells = [];
  for (var r = startRow; r <= endRow; r++) {
    for (var c = startCol; c <= endCol; c++) {
      cells.push({ row: r, col: c });
    }
  }

  return {
    startRow: startRow,
    startCol: startCol,
    numRows: numRows,
    numCols: numCols,
    cells: cells
  };
}

function parseCell(cellStr) {
  var match = cellStr.match(/^([A-Z]+)?(\d+)?$/);
  if (!match) {
    throw new Error('Invalid range format: ' + cellStr);
  }

  var colStr = match[1] || '';
  var rowStr = match[2] || '';

  var col = 0;
  if (colStr) {
    for (var i = 0; i < colStr.length; i++) {
      col = col * 26 + (colStr.charCodeAt(i) - 64);
    }
  }
  var row = rowStr ? parseInt(rowStr) : 0;

  return { row: row, col: col };
}

function parseColumns(colStr) {
  colStr = colStr.toUpperCase().trim();
  var parts = colStr.split(':');
  var startColStr = parts[0];
  var endColStr = parts.length > 1 ? parts[1] : startColStr;

  var startCol = columnToIndex(startColStr);
  var endCol = columnToIndex(endColStr);

  var cols = [];
  for (var i = Math.min(startCol, endCol); i <= Math.max(startCol, endCol); i++) {
    cols.push(i);
  }
  return cols;
}

function columnToIndex(colStr) {
  var col = 0;
  for (var i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  return col;
}

function GPT_SUMMARIZE(text, style = "academic", model = "gemini-2.0-pro", custom = "") {
  if (!text) {
    return "Error: Text is required";
  }
  if (!ALLOWED_STYLES.includes(style)) {
    return `Error: Invalid style. Allowed styles: ${ALLOWED_STYLES.join(", ")}`;
  }
  if (!ALLOWED_MODELS.includes(model)) {
    return `Error: Invalid model. Allowed models: ${ALLOWED_MODELS.join(", ")}`;
  }
  return callAPI(text, style, model, custom);
}