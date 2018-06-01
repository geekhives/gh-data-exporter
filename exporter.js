import { workbook, generate } from './excel';


generate(
    {
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
    },
    workbook({ worksheetTitle: 'My Worksheet', filename: 'sample.xlsx' })
).then(data => {
    console.log(data);
});