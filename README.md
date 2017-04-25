# excelMatlab
This provides a simple interface for writing Microsoft Excel files from MATLAB.

## *Why not just use xlswrite()?*
When called more than ~3 times, xlswrite() is an expensive function in terms of time. excelMatlab provides a simple, optimized approach to writing Excel files when multiple calls are needed. Moreover, xlswrite() requires a string (such as 'A1:B3') representing the range of cells to write. Programmatically this is inconvenient. excelMatlab requests the row and column number of the upper left cell from which to write.

## Examples
Open file called *fileName* and write random matrix to sheet called *sheetName*. The matrix will be anchored to its top left corner, specified by row **1** and column **5** (cell *E1*).
```matlab
data = rand(20);
fullPathToFile = [pwd(), filesep(), 'fileName'];
myExcel = ExcelMatlab(fullPathToFile, 'w');
myExcel.writeToSheet(data, 'sheetName', 1, 5);
myExcel.save();
```

Open file called *fileName* and read cell *G8* from sheet called *sheetName*.
```matlab
fullPathToFile = [pwd(), filesep(), 'fileName'];
myExcel = ExcelMatlab(fullPathToFile);
cellRead = myExcel.readCell('sheetName', 8, 7);
```
