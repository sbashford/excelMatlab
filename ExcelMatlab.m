classdef ExcelMatlab < handle
    properties (Access = 'protected')
        excelApplication
        excelFileWorkbook
        excelFileSheets
        pathToFile
    end
    
    methods
        function self = ExcelMatlab(file)
            stack = dbstack('-completenames', 1);
            if size(stack)
                self.pathToFile = [stack(1).file, '\', file];
            else
                self.pathToFile = [pwd(), '\', file];
            end
            
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
            sheetNumber = self.findSheetNumber(sheet);
            if sheetNumber
                sheetToWrite = self.excelFileSheets(sheetNumber);
            else
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
    
    methods (Access = 'protected')
        function sheetNumber = findSheetNumber(self, sheet)
            numberOfSheets = self.excelFileSheets.Count;
            namesOfSheets = cell(1, numberOfSheets);
            for i = 1:numberOfSheets;
                namesOfSheets{i} = self.excelFileSheets.Item(i).Name;
            end

            [~, sheetNumber] = ismember(sheet, namesOfSheets);
        end
    end
end

