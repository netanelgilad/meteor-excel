var XLSX = Npm.require('xlsx');
var XLS = Npm.require('xlsjs');

ExcelUtils = function (fileType) {
  this.fileType = fileType;
};

ExcelUtils.prototype.sheet_to_json = function (worksheet) {
  if (this.fileType === 'xlsx') {
    return XLSX.utils.sheet_to_json(worksheet);
  }
  else if (this.fileType === 'xls') {
    return XLS.utils.sheet_to_json(worksheet);
  }
};

ExcelUtils.prototype.sheet_to_csv = function (worksheet) {
  if (this.fileType === 'xlsx') {
    return XLSX.utils.sheet_to_csv(worksheet);
  }
  else if (this.fileType === 'xls') {
    return XLS.utils.sheet_to_csv(worksheet);
  }
};

ExcelUtils.prototype.encode_cell = function (cellAddress) {
  if (this.fileType === 'xlsx') {
    return XLSX.utils.encode_cell(cellAddress);
  }
  else if (this.fileType === 'xls') {
    return XLS.utils.encode_cell(cellAddress);
  }
};

ExcelUtils.prototype.encode_range = function (range) {
  if (this.fileType === 'xlsx') {
    return XLSX.utils.encode_range(range);
  }
  else if (this.fileType === 'xls') {
    return XLS.utils.encode_range(range);
  }
};
