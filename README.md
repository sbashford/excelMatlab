# excelMatlab
This provides a simple interface for writing Microsoft Excel files from MATLAB.

## *Why not just use xlswrite()?*
When called more than ~3 times, xlswrite() is an expensive function in terms of time. excelMatlab provides a simple, optimized approach to writing Excel files. Moreover, xlswrite() requires a string (such as 'A1:B3') representing the range of cells to write. Programmatically this is inconvenient. excelMatlab requests the row and column number of the upper left cell from which to write.

## Examples
Open file called *fileName* and write random matrix to sheet called *sheetName*. The matrix will be anchored to its top left corner, specified by row **1** and column **5** (cell *E1*). Once the calling function completes (or the instance *myExcel* is destroyed) the file is saved.
```matlab
data = rand(20);
fullPathToFile = [pwd(), '\fileName.ext'];
myExcel = ExcelMatlab(fullPathToFile);
myExcel.writeToSheet('sheetName', data, 1, 5);
```
