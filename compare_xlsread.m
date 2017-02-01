file = 'thisIsABetterName.xlsx';
sheet = 'thisSheet';
for iterations = 1:10
    tic();
    for i = 1:iterations
        xlsread(file, sheet, 'A1:A10');
    end
    elapsedTime = toc();
    fprintf('Finished xlsread() %d times with %d seconds\n', iterations, elapsedTime);
    tic();
    myExcel = ExcelMatlab(which(file));
    for i = 1:iterations
        myExcel.readNumericColumnRange(sheet, 1, 1, 10);
    end
    delete(myExcel);
    elapsedTime = toc();
    fprintf('Finished excelMatlab %d times with %d seconds\n', iterations, elapsedTime);
end