classdef ExcelMatlab < handle
    properties (Access = private, Constant)
        excelCOM = ExcelCOM()
    end
    
    properties
        workbook
        workbookSheets
        fullPathToFile
        writePermission = false
    end
    
    methods
        function self = ExcelMatlab(varargin)
            assert( ...
                nargin == 1 || nargin == 2, ...
                'ExcelMatlab:invalidNumberArgs', ...
                'Argument error.');
            if nargin == 2
                assert( ...
                    strcmpi(varargin{2}, 'w'), ...
                    'ExcelMatlab:invalidArgument', ...
                    'If seeking write permission, use ''w'' or ''W''.');
            end
            
            fullPathToFile = varargin{1};
            assert( ...
                ischar(fullPathToFile), ...
                'ExcelMatlab:invalidPath', ...
                'Path must be a string.');

            if nargin == 1
                workbook = self.openWorkbookForReading(fullPathToFile);
            else
                workbook = self.openWorkbookForWriting(fullPathToFile);
                self.confirmWritableFile(workbook, fullPathToFile);
            end
            self.workbook = workbook;
            self.workbookSheets = workbook.Sheets;
            self.fullPathToFile = fullPathToFile;
        end
        
        function delete(self)
            if ~isempty(self.workbook)
                self.workbook.Close();
            end
        end
    end

    methods (Access = private)
        function workbook = openWorkbookForWriting(self, fullPathToFile)
            try
                workbook = self.excelCOM.Workbooks.Open(fullPathToFile);
            catch
                workbook = self.excelCOM.Workbooks.Add();
            end
            
            assert( ...
                ~strcmpi(workbook.FileFormat, 'xlCurrentPlatformText'), ...
                'ExcelMatlab:invalidFileFormat', ...
                'The specified file is not a valid excel format.');
        end
        
        function workbook = openWorkbookForReading(self, fullPathToFile)
            try
                workbook = self.excelCOM.Workbooks.Open(fullPathToFile, [], true);
            catch
                error( ...
                    'ExcelMatlab:openFileForReading', ...
                    'unable to read from %s\n', ...
                    fullPathToFile);
            end
            
            assert( ...
                ~strcmpi(workbook.FileFormat, 'xlCurrentPlatformText'), ...
                'ExcelMatlab:invalidFileFormat', ...
                'The specified file is not a valid excel format.');
        end
        
        function confirmWritableFile(self, workbook, fullPathToFile)
            try
                workbook.SaveAs(fullPathToFile);
                self.writePermission = true;
            catch
                error( ...
                    'ExcelMatlab:invalidPath', ...
                    'unable to write to %s\n', ...
                    fullPathToFile);
            end
        end
    end
    
    methods
        function writeToSheet(self, data, sheetName, topLeftRow, topLeftCol)
            assert( ...
                self.writePermission, ...
                'ExcelMatlab:invalidPermission', ...
                'Cannot write to file with current permission.');
            assert( ...
                ischar(sheetName), ...
                'ExcelMatlab:invalidSheetName', ...
                'Sheet must be a string');
            assert( ...
                self.isNonnegativeInteger(topLeftRow) ...
                && self.isNonnegativeInteger(topLeftCol), ...
                'ExcelMatlab:invalidRowCol', ...
                'Row and column must be nonnegative integers.');
            
            bottomRightCol = size(data, 2) + topLeftCol - 1;
            bottomRightRow = size(data, 1) + topLeftRow - 1;
            rangeName = ExcelMatlab.getRangeName( ...
                topLeftCol, ...
                bottomRightCol, ...
                topLeftRow, ...
                bottomRightRow);
            
            sheetToWrite = self.getSheetToWrite(sheetName);
            self.tryWritingToSheet(data, sheetToWrite, rangeName);
            self.workbook.Save();
        end
        
        function cell = readCell(self, sheetName, row, col)
            assert( ...
                ischar(sheetName), ...
                'ExcelMatlab:invalidSheetName', ...
                'Sheet must be a string');
            assert( ...
                self.isNonnegativeInteger(row) ...
                && self.isNonnegativeInteger(col), ...
                'ExcelMatlab:invalidRowCol', ...
                'Row and column must be nonnegative integers.');
            
            rangeName = ExcelMatlab.getRangeName(col, col, row, row);
            
            sheetToRead = self.getSheetToRead(sheetName);
            cell = self.tryReadingFromSheet(sheetToRead, rangeName);
        end
        
        function columnData = readNumericColumnRange(self, sheetName, col, firstRow, lastRow)
            assert( ...
                ischar(sheetName), ...
                'ExcelMatlab:invalidSheetName', ...
                'Sheet must be a string');
            assert( ...
                self.isNonnegativeInteger(col) ...
                && self.isNonnegativeInteger(firstRow) ...
                && self.isNonnegativeInteger(lastRow), ...
                'ExcelMatlab:invalidRowCol', ...
                'Row and column must be nonnegative integers.');
            
            rangeName = ExcelMatlab.getRangeName(col, col, firstRow, lastRow);
            
            sheetToRead = self.getSheetToRead(sheetName);
            columnCell = self.tryReadingFromSheet(sheetToRead, rangeName);
            columnData = cell2mat(columnCell);
        end
        
        function sheetNames = getSheetNames(self)
            numberOfSheets = self.workbookSheets.Count;
            sheetNames = cell(1, numberOfSheets);
            for i = 1:numberOfSheets
                sheetNames{i} = self.workbookSheets.Item(i).Name;
            end
        end
    end
    
    methods (Access = private, Static)
        function is = isNonnegativeInteger(n)
            try
                assert(~isempty(n), 'ExcelMatlab:invalid', '');
                zeros(1, n);
                is = n > 0;
            catch
                is = false;
            end
        end
    end
    
    methods (Access = private)
        function sheetToWrite = getSheetToWrite(self, sheetName)
            sheetNumber = self.findSheetNumber(sheetName);
            if sheetNumber
                sheetToWrite = self.workbookSheets.Item(sheetNumber);
            else
                sheetToWrite = self.addNewSheet(sheetName);
            end
        end
        
        function sheetToRead = getSheetToRead(self, sheetName)
            sheetNumber = self.findSheetNumber(sheetName);
            if sheetNumber
                sheetToRead = self.workbookSheets.Item(sheetNumber);
            else
                error( ...
                    'ExcelMatlab:invalidSheet', ...
                    'Sheet specified not found in workbook.');
            end
        end
        
        function sheetNumber = findSheetNumber(self, sheetName)
            numberOfSheets = self.workbookSheets.Count;
            namesOfSheets = cell(1, numberOfSheets);
            for i = 1:numberOfSheets
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
            rangeName = [ ...
                firstColumnName, ...
                num2str(firstRow), ...
                ':', ...
                lastColumnName, ...
                num2str(lastRow)];
        end
    end
    
    methods (Static, Access = private)
        function columnName = getColumnNameFromNumber(n)
            numberOfLettersInAlphabet = 26;
            if n > numberOfLettersInAlphabet
                offset = floor((n - 1) / numberOfLettersInAlphabet);
                firstLetter = char('A' + offset - 1);
                secondLetter = char('A' + mod(n - 1, numberOfLettersInAlphabet));
                columnName = [firstLetter, secondLetter];
            else
                columnName = char('A' + n - 1);
            end
        end
    end
    
    methods (Access = private)
        function tryWritingToSheet(~, data, sheetToWrite, excelRangeName)
            try
                rangeToWrite = get(sheetToWrite, 'Range', excelRangeName);
                rangeToWrite.Value = data;
            catch
                error('ExcelMatlab:invalidRange', 'Invalid range specified');
            end
        end
        
        function data = tryReadingFromSheet(~, sheetToRead, excelRangeName)
            try
                rangeToRead = get(sheetToRead, 'Range', excelRangeName);
                data = rangeToRead.Value;
            catch 
                error('ExcelMatlab:invalidRange', 'Invalid range specified');
            end
        end
    end
end