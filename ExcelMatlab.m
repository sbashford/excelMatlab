classdef ExcelMatlab < handle
    properties (Access = 'private')
        app
        workbook
        workbookSheets
        fullPathToFile
        successSaving = false
    end
    
    methods
        function self = ExcelMatlab(fullPathToFile)
            assert(ischar(fullPathToFile), 'Path must be a string.');
            assert(~isempty(fileparts(fullPathToFile)), 'Invalid path entered.');
            self.fullPathToFile = fullPathToFile;
            self.startExcel();
            self.openWorkbook();
            self.workbookSheets = self.workbook.Sheets;
            self.confirmWritableFile();
        end
    end

    methods (Access = 'private')      
        function startExcel(self)
            self.app = COM.Excel_Application('server', '', 'IDispatch');
            self.app.DisplayAlerts = false;
        end
        
        function openWorkbook(self)
            try
                self.workbook = self.app.Workbooks.Open(self.fullPathToFile);
            catch
                self.workbook = self.app.Workbooks.Add();
            end
        end
        
        function confirmWritableFile(self)
            try
                self.workbook.SaveAs(self.fullPathToFile);
                self.successSaving = true;
            catch MException
                fprintf(['unable to write to ', self.fullPathToFile, '\n']);
                throw(MException);
            end
        end
    end
    
    methods
        function writeToSheet(self, data, sheetName, topLeftRow, topLeftCol)
            assert(ischar(sheetName), 'Sheet must be a string');
            assert(self.isNonnegativeInteger(topLeftRow) && ...
                   self.isNonnegativeInteger(topLeftCol), 'Row and column must be nonnegative integers.');

            sheetNumber = self.findSheetNumber(sheetName);
            if sheetNumber
                sheetToWrite = self.workbookSheets.Item(sheetNumber);
            else
                sheetToWrite = self.addNewSheet(sheetName);
            end
            excelRangeName = ExcelMatlab.getExcelRangeName(topLeftRow, topLeftCol, data);
            self.tryWritingToSheet(data, sheetToWrite, excelRangeName);
        end
    end
    
    methods (Access = 'private', Static)
        function validity = isNonnegativeInteger(n)
            try
                zeros(1, n);
                validity = n > 0;
            catch
                validity = false;
            end
        end
    end
    
    methods (Access = 'private')
        function sheetNumber = findSheetNumber(self, sheetName)
            numberOfSheets = self.workbookSheets.Count;
            namesOfSheets = cell(1, numberOfSheets);
            for i = 1:numberOfSheets;
                namesOfSheets{i} = self.workbookSheets.Item(i).Name;
            end
            [~, sheetNumber] = ismember(sheetName, namesOfSheets);
        end
        
        function newSheet = addNewSheet(self, sheetName)
            numberOfSheets = self.workbookSheets.Count;
            self.workbookSheets.Add([], self.workbookSheets.Item(numberOfSheets));
            numberOfSheets = numberOfSheets + 1;
            newSheet = self.workbookSheets.Item(numberOfSheets);
            newSheet.Name = sheetName;
        end
    end
    
    methods (Static)
        function rangeName = getExcelRangeName(topLeftRow, topLeftCol, data)
            BottomRightCol = size(data, 2) + topLeftCol - 1;
            BottomRightRow = size(data, 1) + topLeftRow - 1;
            rangeName = ExcelMatlab.getRangeName(topLeftCol, BottomRightCol, topLeftRow, BottomRightRow);
        end
    end
    
    methods (Static, Access = 'private')
        function rangeName = getRangeName(firstColumn, lastColumn, firstRow, lastRow)
            firstColumnName = ExcelMatlab.getColumnNameFromColumnNumber(firstColumn);
            lastColumnName = ExcelMatlab.getColumnNameFromColumnNumber(lastColumn);
            rangeName = [firstColumnName, num2str(firstRow), ':', lastColumnName, num2str(lastRow)];
        end
        
        function columnName = getColumnNameFromColumnNumber(columnNumber)
            numberOfLettersInAlphabet = 26;
            if columnNumber > numberOfLettersInAlphabet
                counter = 0;
                while columnNumber - numberOfLettersInAlphabet > 0
                    columnNumber = columnNumber - numberOfLettersInAlphabet;
                    counter = counter + 1;
                end
                columnName = [char('A' + counter - 1), char('A' + columnNumber - 1)];
            else
                columnName = char('A' + columnNumber - 1);
            end
        end
    end
    
    methods (Access = 'private')
        function tryWritingToSheet(~, data, sheetToWrite, excelRangeName)
            try
                rangeToWrite = get(sheetToWrite, 'Range', excelRangeName);
                rangeToWrite.Value = data;
            catch MException
                fprintf('Invalid range specified\n');
                throw(MException);
            end
        end
    end
    
    methods
        function delete(self)
            if self.successSaving
                self.workbook.SaveAs(self.fullPathToFile);
            end
            Quit(self.app);
            delete(self.app);
        end
    end
end

