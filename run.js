const Splitter = require('./src');

const readXlsxFile = require('read-excel-file/node')
const createExcel = require('excel4node');
var workbook = new createExcel.Workbook();
var worksheet = workbook.addWorksheet('Sheet 1');
let fileInputName = ['./src/input.xlsx','./src/output.xlsx'];
const readBuffer = async () => {
    let rowss
    try {
        await readXlsxFile(fileInputName[0]).then((rows) => {
            rowss = rows;
        });
        return rowss
    } catch (error) {
        console.error(error)
    }
}
function isNotNull(param) {
    if (param === null) {
        return false;
    } else {
        return true;
    }
}
const process = async () => {
    var style = workbook.createStyle({
        font: {
            color: '#000000',
            size: 12
        },
        numberFormat: '$#,##0.00; ($#,##0.00); -'
    });
    var wb = new createExcel.Workbook();
    var ws = wb.addWorksheet('weather Data');
    let counts = 1;
    const bufferExcel = await readBuffer();

    let headerarr = ['ที่อยู่เดิม', 'ที่อยู่', 'ตำบล/แขวง', 'อำเภอ/เขต', 'จังหวัด', 'รหัสไปรษณีย์'];
    let countsHeader = 1;
    headerarr.forEach(async (val) => {
        ws.cell(counts, countsHeader).string(val.toString()).style(style);
        countsHeader++;
    });
    counts++;
    bufferExcel.forEach(async (col) => {
        if (isNotNull(col[0])) {
            ws.cell(counts, 1).string(col[0].toString()).style(style);
            const result = Splitter.split(col[0]);
            if (typeof result !== 'undefined') {
                if (isNotNull(result.address)) {
                    let temp = result.address.toString();
                    temp = temp.replace('/',temp)
                    temp = temp.replace('กรุงเทพฯ',temp)

                    ws.cell(counts, 2).string(result.address.toString()).style(style);
                }
                if (isNotNull(result.subdistrict)) {
                    ws.cell(counts, 3).string(result.subdistrict.toString()).style(style);
                }
                if (isNotNull(result.district)) {
                    ws.cell(counts, 4).string(result.district.toString()).style(style);
                }
                if (isNotNull(result.province)) {
                    ws.cell(counts, 5).string(result.province.toString()).style(style);
                }
                if (isNotNull(result.address)) {
                    ws.cell(counts, 6).string(result.zipcode.toString()).style(style);
                }
           }
        }
        counts++;
    });

    await wb.write(fileInputName[1]);
}
process().then(async () => {
    console.log("success");
});


