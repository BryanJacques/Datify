# Datify
Google Apps Script library for working with tabular data in Sheets. The library creates two custom classes, datarray and datarange. [Datarray](https://github.com/BryanJacques/Datify/blob/main/Datarray.md) is an object for working with two dimensional arrays of data ([[header1, header2],[rowVal1,rowVal2]]), and the [Datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md) object is for working with tabular data in a named range once it has been placed into a Google sheet.  

To use Datify, either copy Datify code from Datify.js into your own apps script file, or authorize and connect to actual library [here](https://script.google.com/home/projects/1VYa1E1zFyi2K3mv4plukIeSsOs2AFEVFlPbXUx7yMWA7Wf3TJEAY7V0l/edit).  

To see an example workflow or run some test code for Datify, see [this](https://docs.google.com/spreadsheets/d/1yqm2cw7Ns-agq43lpXYEdRSodg-8CDXqMvIHokN61J8/edit#gid=2137455568) spreadsheet.

## Datarray

[Datarray](https://github.com/BryanJacques/Datify/blob/main/Datarray.md) function accepts two dimensional array in following format:
```
  [
    ['col 1 header','col 2 header','col 3 header]
    ,['row 1 col 1','row 1 col 2','row 1 col 3']
    ,['row 2 col 1','row 2 col 2','row 2 col 3']
  ]
```
and it returns an object with methods to work with tabular data.

## Datarange

datarange function accepts the name of a named range as input and returns an object with methods to work with data placed on a Google sheet.


