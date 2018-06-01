'use strict';

Object.defineProperty(exports, "__esModule", {
    value: true
});
exports.generate = exports.workbook = undefined;

var _exceljs = require('exceljs');

var _exceljs2 = _interopRequireDefault(_exceljs);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var options = {
    filename: './streamed-workbook.xlsx',
    useStyles: true,
    useSharedStrings: true
};

var workbook = function workbook(_ref) {
    var worksheetTitle = _ref.worksheetTitle,
        filename = _ref.filename;

    var newOptions = Object.assign({}, options, { filename: filename });
    var wb = new _exceljs2.default.stream.xlsx.WorkbookWriter(newOptions);
    wb.addWorksheet(worksheetTitle, {
        pageSetup: {
            paperSize: 9,
            orientation: 'protrait',
            fitToPage: true
        }
    });
    return {
        workbook: wb,
        worksheet: wb.getWorksheet(worksheetTitle),
        options: newOptions
    };
};

var setColumn = function setColumn(_ref2, ws) {
    var colCell = _ref2.colCell,
        rowCell = _ref2.rowCell,
        value = _ref2.value,
        width = _ref2.width;

    var cell = '' + colCell + rowCell;
    ws.getColumn(colCell).width = width;
    ws.getCell(cell).value = value;
};

var generateColumn = function generateColumn(data, ws) {
    try {
        var columns = data.columns;

        if (columns.length > 0) {
            columns.forEach(function (item) {
                setColumn(item, ws);
            });
        }
    } catch (error) {
        throw new Error(error);
    }
};

var generateRows = function generateRows(data, ws) {
    try {
        var rows = data.rows,
            columns = data.columns;

        if (rows.length > 0) {
            rows.forEach(function (item) {
                ws.addRow();
                var row = ws.lastRow;
                columns.forEach(function (col, idx) {
                    var cell = idx + 1;
                    row.getCell(cell).value = item[col.key];
                    row.getCell(cell).alignment = item[col.alignment];
                });
            });
        }
    } catch (error) {
        throw new Error(error);
    }
};

var generate = function generate(data, wb) {
    return new Promise(function (resolve, reject) {
        try {
            generateColumn(data, wb.worksheet);
            generateRows(data, wb.worksheet);
            wb.workbook.commit();
            resolve(wb.options);
        } catch (error) {
            throw new Error(error);
            reject(error);
        }
    });
};

exports.workbook = workbook;
exports.generate = generate;
