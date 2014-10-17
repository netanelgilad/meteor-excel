var XLSX = Npm.require('xlsx');

var XLS = Npm.require('xlsjs');

Excel = function(fileType) {
  if (fileType != 'xlsx' && fileType != 'xls') {
    throw new Meteor.Error(400, "File must be of type xlsx or xls");
  }

  this.fileType = fileType;
};

Excel.prototype.readFile = function(fileName, read_opts) {
  if (this.fileType === 'xlsx') {
    return XLSX.readFile(fileName, read_opts);
  }
  else if (this.fileType == 'xls') {
    return XLS.readFile(fileName, read_opts);
  }
};

Excel.prototype.sheet_to_json = function(worksheet) {
  if (this.fileType === 'xlsx') {
    return XLSX.utils.sheet_to_json(worksheet);
  }
  else if (this.fileType == 'xls') {
    return XLS.utils.sheet_to_json(worksheet);
  }
};