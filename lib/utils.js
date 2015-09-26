var XLSX = Npm.require('xlsx');
var XLS = Npm.require('xlsjs');

ExcelUtils = function (fileType) {
  this.fileType = fileType;
};

ExcelUtils.prototype.sheet_to_json = function (worksheet, options) {
  
  // Ensure backwards compatibility - options could be undefined. 
  var opts = options ? options : {};
  
  if (this.fileType === 'xlsx') {
    return XLSX.utils.sheet_to_json(worksheet, opts);
  }
  else if (this.fileType === 'xls') {
    return XLS.utils.sheet_to_json(worksheet, opts);
  }
};

ExcelUtils.prototype.sheet_to_csv = function (worksheet, options) {
  
  // Ensure backwards compatibility - options could be undefined. 
  var opts = options ? options : {};
  
  if (this.fileType === 'xlsx') {
    return XLSX.utils.sheet_to_csv(worksheet, opts);
  }
  else if (this.fileType === 'xls') {
    return XLS.utils.sheet_to_csv(worksheet, opts);
  }
};

ExcelUtils.prototype.encode_cell = function (cellAddress, options) {
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
