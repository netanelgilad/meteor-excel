meteor-excel
============

___Mainly this package is a wrap for the npm packages listed below. So checkout their documentation to see how to work with excel files properly.___

Parsing and generating excel files (xlsx, xls).

This package uses the npm packages:
* [SheetJS/js-xlsx](https://github.com/SheetJS/js-xlsx)
* [SheetJS/js-xls](https://github.com/SheetJS/js-xls)

## Getting Started
There is nothing like a good example to get started with this package. Check out this [MeteorPad](http://meteorpad.com/pad/2hjNqmwHjDvkxvLC5/Leaderboard) for an exportable leaderboard to excel file. Follow the comments in the meteorpad, mainly in the `server/app.js` file.

## Three steps to use it (basic cover)

1.- GET YOU APP PATH..

```javascript
var fs = Npm.require('fs');
var path = Npm.require('path');
var basepath = path.resolve('.').split('.meteor')[0];
```

2.- CREATE A NEW EXCEL OBJECT

___ This package exports the Excel object to the server. Meaning this package is currently available only to the server side.___

To work with excel files first create an `Excel` object matching the excel file type you want to handle: ___xlsx/xls___. To do that just use:

```javascript
var excel = new Excel('xls');
```

or

```javascript
var excel = new Excel('xlsx');
```

3.- READING XLS/X 

* Read a file

```javascript
var workbook = excel.readFile( basepath+'yourFilesFoler/someExcelFile.xls'); 
```
* Get the SheetNames (this is important to use most of the functions)

```javascript
var yourSheetsName = workbook.SheetNames;
```

* Get a cell

```javascript
console.log("Get the 1st Sheet Name (remember is an array): " + workbook.SheetNames[0]);
console.log("Get Some Cell from it: " + workbook.Sheets[yourSheetsName[0]][
'C37'].v);
```

* Make a JSON out of your excel

```javascript
// We want JSON for this sheet:
var sheet = workbook.Sheets[yourSheetsName[0]]

// You can get the sheet as list of lists.
var options = { header : 1 }

// Or you  can get an object with column headers as keys.  
var options = { header : ['title', 'fName', 'sName' ,'address' ] }

// If options is empty or omitted, it should use the first-row headers by default. 
// However this doesn't seem to work with all Excel worksheets. 
var options = {}

// Generate the JSON like so:
var workbookJson = excel.utils.sheet_to_json( sheet, options );
```

* Make a CSV out of your excel

```javascript
var workbookCsv = excel.utils.sheet_to_csv(workbook.Sheets[yourSheetsName[0]]);
console.log(workbookCsv.length);
```


## Contribute
Used this package? got an example to show? conact me or PR the README and i'll happly add it :)
