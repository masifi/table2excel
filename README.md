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
Create and object `obj` of `Table2Eacel(element, options)` then call `obj.export()`

### Config
- `element` is an element `id` or `class`.
- `option` is a configuration that includes:
  - `exlude` a given `id` and `class` name is skiped during the export.
  - `filename` is a exported file name.  

see the below example that exports a `<table id="tbl">`
```
let t2e = new Table2Excel('#tbl', {
  exclude: "",
  filename: 'file-name'
});
t2e.export();
```
