import Excel from 'exceljs';

const options = {
    filename: './streamed-workbook.xlsx',
    useStyles: true,
    useSharedStrings: true
}

const workbook = ({ worksheetTitle, filename }) => {
    const newOptions = Object.assign({}, options, { filename })
    const wb = new Excel.stream.xlsx.WorkbookWriter(newOptions);
    wb.addWorksheet(worksheetTitle, {
        pageSetup: {
            paperSize: 9, 
            orientation:'protrait',
            fitToPage: true
        }
    });
    return {
        workbook: wb,
        worksheet: wb.getWorksheet(worksheetTitle)
    }
}

const setColumn = ({ colCell, rowCell, value, width }, ws) => {
    const cell = `${colCell}${rowCell}`;
    ws.getColumn(colCell).width = width;
    ws.getCell(cell).value = value;
}

const generateColumn = (columns, ws) => {
    if(columns.length > 0) {
        columns.forEach(item => {
            setColumn(item, ws);
        })
    }
}

const generateRows = (rows, columns, ws) => {
    if(rows.length > 0) {
        rows.forEach(item => {
            ws.addRow();
            const row = ws.lastRow;
            columns.forEach((col, idx) => {
                const cell = idx + 1;
                row.getCell(cell).value = item[col.key];
                row.getCell(cell).alignment = item[col.alignment];
            })
        })
    }
}

export {
    workbook,
    setColumn,
    generateColumn,
    generateRows
}