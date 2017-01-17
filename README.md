# excelMatlab

## Examples

Open file called *fileName* and write random matrix to sheet called *sheetName*.
The matrix will be anchored to its top left corner, specified by row **1** and column **5** (E1).
```matlab
data = rand(20);
myExcel = ExcelMatlab('fileName');
myExcel.writeToSheet('sheetName', data, 1, 5);
```
