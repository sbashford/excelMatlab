file = 'thisIsABetterName.xlsx';
sheet = 'thisSheet';
data = rand(15);
for iterations = 1:10
    tic();
    for i = 1:iterations
        xlswrite(file, data, sheet);
    end
    elapsedTime = toc();
    fprintf('Finished xlswrite() %d time(s) in %d seconds\n', iterations, elapsedTime);
    tic();
    myExcel = ExcelMatlab(which(file), 'w');
    for i = 1:iterations
        myExcel.writeToSheet(data, sheet, 1, 1);
    end
    delete(myExcel);
    elapsedTime = toc();
    fprintf('Finished excelMatlab writeToSheet() %d time(s) in %d seconds\n', iterations, elapsedTime);
end