file = 'thisIsABetterName.xlsx';
sheet = 'thisSheet';
data = rand(15);
for iterations = 1:20
    tic();
    for i = 1:iterations
        xlswrite(file, data, sheet);
    end
    elapsedTime = toc();
    fprintf('Finished xlswrite() %d times with %d seconds\n', iterations, elapsedTime);
    tic();
    myExcel = ExcelMatlab(which(file));
    for i = 1:iterations
        myExcel.writeToSheet(data, sheet, 1, 1);
    end
    delete(myExcel);
    elapsedTime = toc();
    fprintf('Finished excelMatlab %d times with %d seconds\n', iterations, elapsedTime);
end