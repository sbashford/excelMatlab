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
            assert(ischar(fullPathToFile), 'ExcelMatlab:invalidPath', 'Path must be a string.');
            assert(~isempty(fileparts(fullPathToFile)), 'ExcelMatlab:invalidPath', 'Invalid path entered.');
            
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
            catch
                error('ExcelMatlab:invalidPath', 'unable to write to %s\n', self.fullPathToFile);
            end
        end
    end
    
    methods
        function writeToSheet(self, data, sheetName, topLeftRow, topLeftCol)
            assert(ischar(sheetName), 'ExcelMatlab:invalidSheetName', 'Sheet must be a string');
            assert(self.isNonnegativeInteger(topLeftRow) && ...
                   self.isNonnegativeInteger(topLeftCol), 'ExcelMatlab:invalidRowCol', ...
                   'Row and column must be nonnegative integers.');
            
            BottomRightCol = size(data, 2) + topLeftCol - 1;
            BottomRightRow = size(data, 1) + topLeftRow - 1;
            rangeName = ExcelMatlab.getRangeName(topLeftCol, BottomRightCol, topLeftRow, BottomRightRow);
            
            sheetToWrite = self.getSheetToWrite(sheetName);
            self.tryWritingToSheet(data, sheetToWrite, rangeName);
        end
        
        function cell = readCell(self, sheetName, row, col)
            assert(ischar(sheetName), 'ExcelMatlab:invalidSheetName', 'Sheet must be a string');
            assert(self.isNonnegativeInteger(row) && ...
                   self.isNonnegativeInteger(col), 'ExcelMatlab:invalidRowCol', ...
                   'Row and column must be nonnegative integers.');
            
            rangeName = ExcelMatlab.getRangeName(col, col, row, row);
            
            sheetToRead = self.getSheetToWrite(sheetName);
            cell = self.tryReadingFromSheet(sheetToRead, rangeName);
        end
        
        function columnData = readNumericColumnRange(self, sheetName, col, firstRow, lastRow)
            assert(ischar(sheetName), 'ExcelMatlab:invalidSheetName', 'Sheet must be a string');
            assert(self.isNonnegativeInteger(col) && ...
                   self.isNonnegativeInteger(firstRow) && ...
                   self.isNonnegativeInteger(lastRow), 'ExcelMatlab:invalidRowCol', ...
                   'Row and column must be nonnegative integers.');
            
            rangeName = ExcelMatlab.getRangeName(col, col, firstRow, lastRow);
            
            sheetToRead = self.getSheetToWrite(sheetName);
            columnCell = self.tryReadingFromSheet(sheetToRead, rangeName);
            columnData = cell2mat(columnCell);
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
        function sheetToWrite = getSheetToWrite(self, sheetName)
            sheetNumber = self.findSheetNumber(sheetName);
            if sheetNumber
                sheetToWrite = self.workbookSheets.Item(sheetNumber);
            else
                sheetToWrite = self.addNewSheet(sheetName);
            end
        end
        
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
        function rangeName = getRangeName(firstColumn, lastColumn, firstRow, lastRow)
            firstColumnName = ExcelMatlab.getColumnNameFromNumber(firstColumn);
            lastColumnName = ExcelMatlab.getColumnNameFromNumber(lastColumn);
            rangeName = [firstColumnName, num2str(firstRow), ':', lastColumnName, num2str(lastRow)];
        end
    end
    
    methods (Static, Access = 'private')
        function columnName = getColumnNameFromNumber(n)
            numberOfLettersInAlphabet = 26;
            if n > numberOfLettersInAlphabet
                offset = floor((n - 1) / numberOfLettersInAlphabet);
                columnName = [char('A' + offset - 1), char('A' + mod(n - 1, numberOfLettersInAlphabet))];
            else
                columnName = char('A' + n - 1);
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
        
        function data = tryReadingFromSheet(~, sheetToRead, excelRangeName)
            try
                rangeToRead = get(sheetToRead, 'Range', excelRangeName);
                data = rangeToRead.Value;
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

