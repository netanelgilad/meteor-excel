Package.describe({
  name: 'netanelgilad:excel',
  summary: 'Parse excel worksheets for your meteor app.',
  version: '0.1.1',
  git: 'https://github.com/netanelgilad/meteor-excel'
});

Npm.depends({
  'xlsx' : '0.7.11',
  'xlsjs' : '0.7.1'
});

Package.onUse(function(api) {
  api.versionsFrom('0.9.0.1');

  api.addFiles('lib/utils.js', 'server');
  api.addFiles('lib/workbook.js', 'server');
  api.addFiles('lib/worksheet.js', 'server');
  api.addFiles('netanelgilad:excel.js', 'server');

  api.export('Excel');
});

Package.onTest(function(api) {
  api.use('tinytest');
  api.use('netanelgilad:excel');
  api.addFiles('netanelgilad:excel-tests.js');
});
