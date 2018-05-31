import { workbook, setColumn, generateColumn, generateRows } from './excel';


const wb = workbook({ worksheetTitle: 'My Worksheet', filename: 'sample.xlsx' });

const data = {
    columns: [
        {
            colCell: 'A', rowCell: 1, value: 'REFERENCE NUMBER', width: 20, key: 'refno',
            alignment: { vertical: 'middle', horizontal: 'center'}
        },
        {
            colCell: 'B', rowCell: 1, value: 'TRANSACTION DATE', width: 20, key: 'date',
            alignment: { vertical: 'middle', horizontal: 'center'}
        },
        {
            colCell: 'C', rowCell: 1, value: 'NAME', width: 20, key: 'name',
            alignment: { vertical: 'middle', horizontal: 'center'}
        }
    ],
    rows: [
        {
            refno: '1111',
            date: '2018-01-01',
            name: 'Paul'
        },
        {
            refno: '2222',
            date: '2018-01-01',
            name: 'Gerhard'
        },
        {
            refno: '2222',
            date: '2018-01-01',
            name: 'Jinggo'
        },
        {
            refno: 'gerhard',
            date: '2018-01-01',
            name: 'Rommel'
        }
    ]
}

generateColumn(data, wb.worksheet);
generateRows(data, wb.worksheet);

wb.worksheet.pageSetup.printArea = 'A1:Q29';
wb.workbook.commit();