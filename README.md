# easySheet Quickstart

easySheet is a mini class to release your workload counting columns in ***Google Spreadsheet*** when coding with ***Google Script***.

Beginers can use ***easySheet*** to script easier by saving time of counting columns one by one.

## How to prepare

If you are editing a Google script, create a new .gs file and paste all lines from easySheet.gs in the ***Script Editor***.

To define the dictionary for easySheet, create a sheet in the Google Spreadsheet in which the code is written. The name of the sheet can be defined by yourself.

The dictionary sheet should be written in the format below.

|          |SeriesName1|SeriesName2 |
|----------|-----------|------------|
|Variable1 |Col.Name1  |Col.Name2   |
|Variable2 |Col.Name1  |Col.Name2   |
|Variable3 |Col.Name1  |Col.Name2   |

- `SeriesName1`: the series name defined by yourself.
- `Variable`: the name of the variable that you use for a specific column name in a specific sheet. No space or special symbol is allowed in variable names. Use ***'\_\_skip\_\_'*** if the sheet does not contain this variable.
- `Col. Name`: the column name refers to the left variable name.

**DO NOT** use these variable names since they are **reserved**: ***idSeriesArray, errorCode, keysMissing, keysDuplicated***

## How to start

Construct function requires a Sheet object (refer to the Google Script documentation for more about class Sheet). For example:

  ```
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('structure');
  var sheetDic = new easySheet(sheet);
  ```

When the above code was run, the constructor will check whether there were duplicated Col.Names in the same SheetName and pop-up a warning if so.


## How to use

Parameters should be defined as below:
- `sheet` a Sheet object (but not the sheet name string), refers to the sheet you are working
- `idSeries` the SeriesName1 you want to apply
- `idRow` or `idCol` the row number or colunm number in which the identifiers lay
- `dataRow` or `dataCol` the row number or colunm in which the data lay


The easySheet class provides four main methods.

- `getNumByRow(sheet,idSeries,idRow)`
This method returns an object. Keys refer to the variable names in the structure sheet. Values are the numbers of columns that refer to a variable in the defined sheet.

- `getNumByCol(sheet,idSeries,idCol)`
This method returns an object. Keys refer to the variable names in the structure sheet. Values are the numbers of rows that refer to a variable in the defined sheet.

- `getDataByRow(sheet,idSeries,idRow,dataRow`
This method returns an object. Keys refer to the variable names in the structure sheet. Values are the values of variable names read from the given specific line number. The method is based on getNumByRow.

- `getDataByCol(sheet,idSeries,idCol,dataCol`
This method returns an object. Keys refer to the variable names in the structure sheet. Values are the values of variable names read from the given specific column number. The method is based on getNumByCol.


***getNumByRow*** and ***getNumByCol*** also contain some defined keys and values:
- `errorCode` int, suggest matching status, 0 for normal, 1 for columns missing, 2 for columns duplicated (more than one columns are using the same title), 3 for both missing and duplicate
- `keysMissing` nullable list, contains all missing columns by its variable name
- `keysDuplicated` nullable list, contains all duplicated columns by its variable name

## Credits

This mini class is developed by ***hirobank***.

Please submit an issue forif you have any advice for ***easySheet***.

**This is the end of the document. Discover more with yourself!**
