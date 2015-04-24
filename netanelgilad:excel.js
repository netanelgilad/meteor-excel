var XLSX = Npm.require('xlsx');
var XLS = Npm.require('xlsjs');

Excel = function (fileType) {
  if (fileType != 'xlsx' && fileType != 'xls') {
    throw new Meteor.Error(400, "File must be of type xlsx or xls");
  }

  this.fileType = fileType;
  this.utils = new ExcelUtils(this.fileType);
};

Excel.prototype.readFile = function (fileName, read_opts) {
  if (this.fileType === 'xlsx') {
    return XLSX.readFile(fileName, read_opts);
  }
  else if (this.fileType == 'xls') {
    return XLS.readFile(fileName, read_opts);
  }
};

Excel.prototype.read = function (file, read_opts) {
  if (this.fileType === 'xlsx') {
    return XLSX.read(file, read_opts);
  }
  else if (this.fileType == 'xls') {
    return XLS.read(file, read_opts);
  }
};

Excel.prototype.createWorkbook = function () {
  return new Workbook(this.fileType);
};

Excel.prototype.createWorksheet = function () {
  return new Worksheet(this.fileType);
};