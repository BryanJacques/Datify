# Datify
Google Apps Script library for working with tabular data in Sheets. The library creates two custom classes, datarray and datarange. [Datarray](https://github.com/BryanJacques/Datify/blob/main/Datarray.md) is an object for working with two dimensional arrays of data ([[header1, header2],[rowVal1,rowVal2]]), and the datarange object is for working with tabular data in a named range once it has been placed into a Google sheet.

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


