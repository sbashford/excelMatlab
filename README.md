# excelMatlab
This provides a simple interface for writing Microsoft Excel files from MATLAB.

## *Why not just use xlswrite()?*
xlswrite() is an expensive function in terms of time. excelMatlab provides a simple, optimized approach to writing Excel files. Often MATLAB programs that write to Excel files invoke xlswrite() multiple times thereby compounding the time cost. Moreover, xlswrite() requires a string (such as 'A1:B3') representing the range of cells to write. Programmatically this is not convenient. excelMatlab requests the row and column number of the upper left cell from which to write.

## Examples
Open file called *fileName* and write random matrix to sheet called *sheetName*.
The matrix will be anchored to its top left corner, specified by row **1** and column **5** (E1).
```matlab
data = rand(20);
myExcel = ExcelMatlab('fileName');
myExcel.writeToSheet('sheetName', data, 1, 5);
```
