var XLSX = Npm.require('xlsx');
var XLS = Npm.require('xlsjs');

Worksheet = function (fileType) {
  this.fileType = fileType;
  this.range = {s: {c: 10000000, r: 10000000}, e: {c: 0, r: 0}};
  this['!merges'] = [];
};

Worksheet.prototype.writeToCell = function (row, col, value) {
  var utils = new ExcelUtils(this.fileType);
  var cellAddress = utils.encode_cell({c: col, r: row});

  var cell = {
    v: value
  };

  if (typeof cell.v === 'number') cell.t = 'n';
  else if (typeof cell.v === 'boolean') cell.t = 'b';
  else if (cell.v instanceof Date) {
    cell.t = 'n';
    cell.z = XLSX.SSF._table[14];
    cell.v = datenum(cell.v);
  }
  else cell.t = 's';

  this[cellAddress] = cell;

  if (this.range.s.r > row) this.range.s.r = row;
  if (this.range.s.c > col) this.range.s.c = col;
  if (this.range.e.r < row) this.range.e.r = row;
  if (this.range.e.c < col) this.range.e.c = col;

  this['!ref'] = utils.encode_range(this.range);
};

Worksheet.prototype.writeObjectArray = function (row, col, objArray, objDefinitions) {
  if (objArray.length === 0)
    return;

  var self = this;

  if (_.isUndefined(objDefinitions)) {
    objDefinitions = {};
    for (var prop in objArray[0]) {
      if (objArray[0].hasOwnProperty(prop)) {
        objDefinitions[prop] = prop;
      }
    }
  }

  var currentHeaderCol = col;
  _.forEach(objDefinitions, function(definition) {
    var header = _.isObject(definition) ? definition.header : definition;
    self.writeToCell(row, currentHeaderCol, header);
    currentHeaderCol++;
  });

  var currentRow = row + 1;
  _.forEach(objArray, function(item) {
    var currentCol = col;
    _.forEach(objDefinitions, function(definition, field) {
      var itemData = _.isUndefined(definition.transform) ? item[field] : definition.transform.apply(item, [item[field]]);
      self.writeToCell(currentRow, currentCol, itemData);
      currentCol++;
    });
    currentRow++;
  });
};

Worksheet.prototype.mergeCells = function(startRow, startCol, endRow, endCol) {
  this['!merges'].push({s:{r:startRow, c:startCol}, e:{r:endRow, c:endCol}});
};

Worksheet.prototype.setColumnProperties = function(columns) {
  this['!cols'] = columns;
};

//////////////////////

function datenum(v, date1904) {
  if (date1904) v += 1462;
  var epoch = Date.parse(v);
  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}