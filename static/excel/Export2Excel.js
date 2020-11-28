/* eslint-disable */
require('script-loader!file-saver');
require('./Blob.js'); //blob.js也是网上找的，下面会贴上代码
require('script-loader!xlsx/dist/xlsx.core.min');    //注意 直接import xlsx-style会报错，因为npm install xlsx-style 下载下来的依赖 源码有错，需要修改，下面会讲到
import XLSX from "xlsx-style"
var  $  = require ('jquery')
function generateArray(table) {
    var out = [];
    var rows = table.querySelectorAll('tr');
    var ranges = [];
    for (var R = 0; R < rows.length; ++R) {
        var outRow = [];
        var row = rows[R];
        var columns = row.querySelectorAll('td');
        for (var C = 0; C < columns.length; ++C) {
            var cell = columns[C];
            var colspan = cell.getAttribute('colspan');
            var rowspan = cell.getAttribute('rowspan');
            var cellValue = cell.innerText;
            if (cellValue !== "" && cellValue == +cellValue) cellValue = +cellValue;

            //Skip ranges
            ranges.forEach(function(range) {
                if (R >= range.s.r && R <= range.e.r && outRow.length >= range.s.c && outRow.length <= range.e.c) {
                    for (var i = 0; i <= range.e.c - range.s.c; ++i) outRow.push(null);
                }
            });

            //Handle Row Span
            if (rowspan || colspan) {
                rowspan = rowspan || 1;
                colspan = colspan || 1;
                ranges.push({
                    s: {
                        r: R,
                        c: outRow.length
                    },
                    e: {
                        r: R + rowspan - 1,
                        c: outRow.length + colspan - 1
                    }
                });
            };

            //Handle Value
            outRow.push(cellValue !== "" ? cellValue : null);

            //Handle Colspan
            if (colspan)
                for (var k = 0; k < colspan - 1; ++k) outRow.push(null);
        }
        out.push(outRow);
    }
    return [out, ranges];
};

function datenum(v, date1904) {
    if (date1904) v += 1462;
    var epoch = Date.parse(v);
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

function sheet_from_array_of_arrays(data, opts) {
    var ws = {};
    var range = {
        s: {
            c: 10000000,
            r: 10000000
        },
        e: {
            c: 0,
            r: 0
        }
    };
    for (var R = 0; R != data.length; ++R) {
        for (var C = 0; C != data[R].length; ++C) {
            if (range.s.r > R) range.s.r = R;
            if (range.s.c > C) range.s.c = C;
            if (range.e.r < R) range.e.r = R;
            if (range.e.c < C) range.e.c = C;
            var cell = {
                v: data[R][C]
            };
            if (cell.v == null) continue;
            var cell_ref = XLSX.utils.encode_cell({
                c: C,
                r: R
            });

            if (typeof cell.v === 'number') cell.t = 'n';
            else if (typeof cell.v === 'boolean') cell.t = 'b';
            else if (cell.v instanceof Date) {
                cell.t = 'n';
                cell.z = XLSX.SSF._table[14];
                cell.v = datenum(cell.v);
            } else cell.t = 's';

            ws[cell_ref] = cell;
        }
    }
    if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
    return ws;
}

function Workbook() {
    if (!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}

function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}

export function export_table_to_excel(id) {
    var theTable = document.getElementById(id);
    console.log('a')
    var oo = generateArray(theTable);
    var ranges = oo[1];

    /* original data */
    var data = oo[0];
    var ws_name = "SheetJS";
    console.log(data);

    var wb = new Workbook(),
        ws = sheet_from_array_of_arrays(data);

    /* add ranges to worksheet */
    // ws['!cols'] = ['apple', 'banan'];              //合并单元格
    ws['!merges'] = ranges;

    /* add worksheet to workbook */
    wb.SheetNames.push(ws_name);
    wb.Sheets[ws_name] = ws;

    var wbout = XLSX.write(wb, {
        bookType: 'xlsx',
        bookSST: false,
        type: 'binary'
    });

    saveAs(new Blob([s2ab(wbout)], {
        type: "application/octet-stream"
    }), "test.xlsx")
}

function formatJson(jsonData) {
    console.log(jsonData)
}
// defaultTitle excel 文件名
// merges 合并单元格
// multiHeader  多表头
// th   单表头
//data  表格数据

export function export_json_to_excel(multiHeader, th, merges, data, defaultTitle) {
    data = [...data];
    data.unshift(th);


    if (multiHeader) {
        data.unshift(multiHeader);
    }
    var ws_name = "SheetJS";

    var wb = new Workbook(),
        ws = sheet_from_array_of_arrays(data);

    /* add worksheet to workbook */
    wb.SheetNames.push(ws_name);        //合并单元格
    ws["!merges"] = merges;
    wb.Sheets[ws_name] = ws;

    var dataInfo = wb.Sheets[wb.SheetNames[0]];
    var cellArr = merges.map(c => c.s);
    var secArr = merges.map(c => c.e);
    var cellArr1 = [];
    cellArr.forEach(cellObj => {
        var cell_ref = XLSX.utils.encode_cell({
            c: cellObj.c,
            r: cellObj.r
        });
        cellArr1.push(cell_ref);
    });
         //设置单元格样式
    const borderAll = {
        border: {
            //单元格外侧框线
            top: {
                style: "thin"
            },
            bottom: {
                style: "thin"
            },
            left: {
                style: "thin"
            },
            right: {
                style: "thin"
            }
        }
    };
    //给所有单元格加上边框
    for (var i in dataInfo) {
        if (i == '!ref' || i == '!merges' || i == '!cols' || $.inArray(i, cellArr1) >= 0) {

        } else {
            dataInfo[i + ''].s = borderAll;
        }
    }
    //设置单元格背景色、字体以及字体大小等
    var bgColArr = ["FDE9D9", 'FFFF00', 'DAEEF3', 'CCC0DA', 'C5D9F1'];
    //设置主标题样式
    var headerStyle = {
        font: {
            name: '宋体',
            color: {
                rgb: "303133"
            },
            bold: true,
            italic: false,
            underline: false
        },
        alignment: {
            horizontal: "center",
            vertical: "center"
        },
        fill: {

        }
    };
    let mm = 0;
    cellArr1.forEach(cellObj => {
        var hStyle = Object.assign({}, headerStyle);
        hStyle.fill = {
            fgColor: {
                rgb: bgColArr[mm]
            }
        };
        dataInfo[cellObj].s = hStyle;
        mm++;
    });

    var secCellStyle = Object.assign({}, borderAll, headerStyle);
    secArr.forEach((s, index) => {
        var hStyle = Object.assign({}, secCellStyle);
        hStyle.fill = {
            fgColor: {
                rgb: bgColArr[index]
            }
        };
        var startIndex = 0;
        if (index == 0) {
            startIndex = 0;
        } else {
            startIndex = secArr[index - 1].c + 1;
        }
        for (var se = startIndex; se <= s.c; se++) {
            var cell_ref = XLSX.utils.encode_cell({
                c: se,
                r: 1
            });
            dataInfo[cell_ref].s = hStyle;
        }
    });

    var wbout = XLSX.write(wb, {
        bookType: 'xlsx',
        bookSST: false,
        type: 'binary'
    });
    var title = defaultTitle || '列表'
    saveAs(new Blob([s2ab(wbout)], {
        type: "application/octet-stream"
    }), title + ".xlsx")
}