'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});

var _xlsx = require('xlsx');

var _xlsx2 = _interopRequireDefault(_xlsx);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

exports.default = function (filename, data) {
  var wb = _xlsx2.default.utils.book_new();
  var columns = data.columns,
      rows = data.rows;


  var newColumns = columns.map(function (item) {
    return item.value;
  });

  var newRows = rows.map(function (item) {
    return Object.values(item);
  });

  var newData = [newColumns].concat(newRows);

  var ws = _xlsx2.default.utils.aoa_to_sheet(newData);
  _xlsx2.default.utils.book_append_sheet(wb, ws, 'Worksheet');
  _xlsx2.default.writeFile(wb, filename);
};
