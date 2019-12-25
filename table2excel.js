class Table2Excel {
    constructor(element, options = {}) {
        this.element = document.querySelectorAll(element);
        this.settings = {};
        this.settings.exclude = (options.exclude !== undefined && options.exclude !== "") ? options.exclude : ".skip";
        this.settings.name = (options.name !== undefined && options.name !== "") ? options.name : "table2excel";
        this.settings.filename = (options.filename !== undefined && options.filename !== "") ? options.filename : "table2excel";
    }
    export () {
        var utf8Heading = "<meta http-equiv=\"content-type\" content=\"application/vnd.ms-excel; charset=UTF-8\">";
        this.template = {
            head: "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns=\"http://www.w3.org/TR/REC-html40\">" + utf8Heading + "<head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets>",
            sheet: {
                head: "<x:ExcelWorksheet><x:Name>",
                tail: "</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>"
            },
            mid: "</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body>",
            table: {
                head: "<table>",
                tail: "</table>"
            },
            foot: "</body></html>"
        };
        this.tableRows = [];

        this.element.forEach((o, i) => {
            var tempRows = "";
            o.querySelectorAll("tr:not(" + this.settings.exclude + ")").forEach((o, i) => {
                tempRows += "<tr>" + o.innerHTML + "</tr>";
            });
            this.tableRows.push(tempRows);
        });
        this.tableToExcel(this.tableRows, this.settings.name, this.settings.sheetName);
    }

    tableToExcel(table, name, sheetName) {
        var fullTemplate = "",
            i, link, a;
        this.uri = "data:application/vnd.ms-excel;base64,";
        this.base64 = function (s) {
            return window.btoa(unescape(encodeURIComponent(s)));
        };
        this.format = function (s, c) {
            return s.replace(/{(\w+)}/g, function (m, p) {
                return c[p];
            });
        };
        sheetName = typeof sheetName === "undefined" ? "Sheet" : sheetName;
        this.ctx = {
            worksheet: name || "Worksheet",
            table: table,
            sheetName: sheetName,
        };
        fullTemplate = this.template.head;
        if (Array.isArray(table)) {
            for (i in table) {
                fullTemplate += this.template.sheet.head + sheetName + i + this.template.sheet.tail;
            }
        }
        fullTemplate += this.template.mid;
        if (Array.isArray(table)) {
            for (i in table) {
                fullTemplate += this.template.table.head + "{table" + i + "}" + this.template.table.tail;
            }
        }
        fullTemplate += this.template.foot;
        for (i in table) {
            this.ctx["table" + i] = table[i];
        }
        delete this.ctx.table;
        if (typeof msie !== "undefined" && msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./)) {
            if (typeof Blob !== "undefined") {
                fullTemplate = [fullTemplate];
                var blob1 = new Blob(fullTemplate, {
                    type: "text/html"
                });
                window.navigator.msSaveBlob(blob1, this.getFileName(this.settings));
            } else {
                txtArea1.document.open("text/html", "replace");
                txtArea1.document.write(this.format(fullTemplate, this.ctx));
                txtArea1.document.close();
                txtArea1.focus();
                sa = txtArea1.document.execCommand("SaveAs", true, this.getFileName(this.settings));
            }
        } else {
            link = this.uri + this.base64(this.format(fullTemplate, this.ctx));
            a = document.createElement("a");
            a.download = this.getFileName(this.settings);
            a.href = link;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        }
        return true;
    }
    getFileName() {
        return (this.settings.filename ? this.settings.filename : "table2excel") + (this.settings.fileext ? this.settings.fileext : ".xls");
    }
}
