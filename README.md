# table2excel
Pure javascript html table to Excel exporter

## Demo
[Exporting HTML table to EXCEL](https://jsfiddle.net/masifi/c3gazmuq)

## Installation and Usage

Installation
------------
Load `table2excel.min.js` in the project header:
```
<script type="text/javascript" src="table2excel.min.js"></script>
```

Usage
-----
Create and object `obj` of `Table2Eacel(element. options)` then call `obj.export()`
see the below usage
```
let t2e = new Table2Excel('#tbl', {
  exclude: "",
  name: "table-name",
  filename: 'file-name'
});
t2e.export();
```
