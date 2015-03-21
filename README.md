meteor-excel
============

Parsing and generating excel files (xlsx, xls).

This package uses the npm packages:
* [SheetJS/js-xlsx](https://github.com/SheetJS/js-xlsx)
* [SheetJS/js-xls](https://github.com/SheetJS/js-xls)

## Getting Started
There is nothing like a good example to get started with this package. Check out this [MeteorPad](http://meteorpad.com/pad/2hjNqmwHjDvkxvLC5/Leaderboard) for an exportable leaderboard to excel file. Follow the comments in the meteorpad, mainly in the `server/app.js` file.

### Using the Excel object

This package exports the `Excel` object to the server. Meaning this package is currently available only to the server side.

To work with excel files first create an `Excel` object matching the excel file type you want to handle: xlsx/xls. To do that just use: `var excel = new Excel('xlsx')`.

Mainly this package is a wrap for the npm packages listed above. So checkout their documentation to see how to work with excel files.

## Contribute
Used this package? got an example to show? conact me or PR the README and i'll happly add it :)
