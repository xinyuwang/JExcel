/*
    Name : JExcel
    Author : Xinyu Wang
    Date : 2016-3-26
    Dependency : xlsx, FileSaver
    Introduction : A simply way to build excel via javascript.
*/

function JExcel() {

}

JExcel.prototype._SheetNames = [];
JExcel.prototype._Sheets = {};
JExcel.prototype._CurrentSheet = "";

JExcel.prototype._Assert = function (expr, msg) {
    if (!expr) {
        throw msg;
    }
}

JExcel.prototype._StringToArrayBuffer = function (s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}

JExcel.prototype._InitSheet = function (sheetName) {
    var _this = this;
    var sheet = {};
    sheet._Name = sheetName;
    sheet._Value = [];
    sheet._Owner = this;

    sheet.Set = function (row, col, val) {
        _this._Assert(row >= 0 && col >= 0, "row and col must above zero.");
        if (!this._Value[row]) {
            this._Value[row] = [];
        }
        this._Value[row][col] = val;

        return this;
    };

    sheet.Get = function (row, col) {
        _this._Assert(row >= 0 && col >= 0, "row and col must above zero.");
        if (this._Value[row]) {
            if (this._Value[row][col]) {
                return this._Value[row][col];
            }
        }

        return null;
    };

    sheet.ToSheet = function () {

        var ws = {};
        var range = { s: { r: 10000000, c: 10000000 }, e: { r: 0, c: 0 } };

        var min = function (a, b) {
            return a < b ? a : b;
        }
        var max = function (a, b) {
            return a > b ? a : b;
        }

        for (var row = 0; row < this._Value.length; row++) {

            if (!this._Value[row]) {
                this._Value[row] = [];
            }

            for (var col = 0; col < this._Value[row].length; col++) {

                range.s.r = min(range.s.r, row);
                range.s.c = min(range.s.c, col);

                range.e.r = max(range.e.r, row);
                range.e.c = max(range.e.c, col);

                var cell = { v: this._Value[row][col] };
                if (cell.v == null) {
                    continue;
                }
                var cell_ref = XLSX.utils.encode_cell({ c: col, r: row });

                if (typeof cell.v === 'number') {
                    cell.t = 'n';
                }
                else if (typeof cell.v === 'boolean') {
                    cell.t = 'b';
                }
                else if (cell.v instanceof Date) {
                    cell.t = 'n';
                    cell.z = XLSX.SSF._table[14];
                    cell.v = datenum(cell.v);
                }
                else {
                    cell.t = 's';
                }

                ws[cell_ref] = cell;
            }
        }

        if (range.s.c < 10000000) {
            ws['!ref'] = XLSX.utils.encode_range(range);
        }

        return ws;
    };

    return sheet;
}

JExcel.prototype.SetSheet = function (sheetName) {
    this._Assert(sheetName != undefined && sheetName != "", "sheetName must be a significant string.");

    if (this._CurrentSheet != sheetName) {

        if (this._Sheets[sheetName] == undefined) {
            this._SheetNames.push(sheetName);
            this._Sheets[sheetName] = this._InitSheet(sheetName);
        }

        this._CurrentSheet = sheetName;

    }

    return this._Sheets[this._CurrentSheet];

}

JExcel.prototype.SaveAs = function (fileName) {

    if (!/.*\.xlsx/.test(fileName)) {
        fileName += ".xlsx";
    }

    var excel = {
        SheetNames: [],
        Sheets: {}
    };

    for (var i = 0; i < this._SheetNames.length; i++) {
        var name = this._SheetNames[i];
        excel.SheetNames.push(name);
        excel.Sheets[name] = this._Sheets[name].ToSheet();
    }

    var excelOut = XLSX.write(excel, {
        bookType: 'xlsx', bookSST: false, type: 'binary'
    });

    /* the saveAs call downloads a file on the local machine */
    saveAs(new Blob([this._StringToArrayBuffer(excelOut)], { type: "" }), fileName);


}
