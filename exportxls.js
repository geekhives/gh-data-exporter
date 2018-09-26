import _ from 'lodash';
import XLSX from 'xlsx';

export default (filename, data) => {
  const wb = XLSX.utils.book_new();
  const { columns, rows } = data;

  const newColumns = columns.map(item => {
    return item.value;
  })

  const newRows = rows.map(item => {
    return _.values(item);
  })

  const newData = _.concat([newColumns], newRows);

  const ws = XLSX.utils.aoa_to_sheet(newData);
  XLSX.utils.book_append_sheet(wb, ws, 'Worksheet');
  XLSX.writeFile(wb, filename);
}