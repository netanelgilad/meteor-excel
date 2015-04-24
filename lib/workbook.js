var XLSX = Npm.require('xlsx');
var XLS = Npm.require('xlsjs');

Workbook = function (fileType) {
  this.fileType = fileType;

  this.SheetNames = [];
  this.Sheets = {};
};

Workbook.prototype.addSheet = function (sheetName, sheet) {
  this.SheetNames.push(sheetName);
  this.Sheets[sheetName] = sheet;
};

Workbook.prototype.writeToFile = function (filePath) {
  if (this.fileType === 'xlsx') {
    return XLSX.writeFile(this, filePath);
  }
  else if (this.fileType == 'xls') {
    return XLS.writeFile(this, filePath);
  }
};
