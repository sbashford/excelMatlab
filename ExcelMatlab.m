classdef ExcelMatlab < handle
    properties (Access = 'protected')
        excelApplication
        excelFileWorkbook
        excelFileSheets
        pathToFile
    end
    
    methods
        function self = ExcelMatlab(file)
            self.pathToFile = file;
            self.excelApplication = actxserver('Excel.Application');
            self.excelApplication.DisplayAlerts = false;
            try
                self.excelFileWorkbook = self.excelApplication.Workbooks.Open(pathToFile); 
            catch
                self.excelFileWorkbook = self.excelApplication.Workbooks.Add();
            end
            self.excelFileSheets = self.excelFileWorkbook.Sheets;
        end
        
        function delete(self)
            try
                self.excelFileWorkbook.SaveAs(self.pathToFile);
            catch
                fprintf('Excel file did not save successfully\n');
            end
            Quit(self.excelApplication);
            delete(self.excelApplication);
        end
        
        function writeToSheet(self, data, sheet, topLeftRow, topLeftCol)
            try
                sheetToWrite = self.excelFileSheets(sheet);
            catch
                numberOfSheets = self.excelFileSheets.Count;
                self.excelFileSheets.Add([], self.excelFileSheets.Item(numberOfSheets));
                numberOfSheets = numberOfSheets + 1;
                sheetToWrite = self.excelFileSheets.Item(numberOfSheets);
                sheetToWrite.Name = sheet;
            end
            
            BottomRightCol = size(data, 2) + topLeftCol - 1;
            BottomRightRow = size(data, 1) + topLeftRow - 1;
            excelRange = getExcelRangeString(topLeftCol, BottomRightCol, topLeftRow, BottomRightRow);
            rangeToWrite = get(sheetToWrite, 'Range', excelRange);
            rangeToWrite.Value = data;
        end
    end
end

