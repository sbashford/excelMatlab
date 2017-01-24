classdef ExcelSheetWriter < handle   
    properties
        workbookSheets
        sheetName
        sheetNumber
    end
    
    methods
        function self = ExcelSheetWriter(workbookSheets, sheetName)
            self.workbookSheets = workbookSheets;
            self.sheetName = sheetName;
            self.sheetNumber = self.findSheetNumber(sheetName);
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
    end
    
    methods
        function write(self, data, topLeftRow, topLeftCol)
            if self.sheetNumber
                sheetToWrite = self.workbookSheets.Item(self.sheetNumber);
            else
                sheetToWrite = self.addNewSheet();
            end
            excelRangeName = ExcelSheetWriter.getExcelRangeName(topLeftRow, topLeftCol, data);
            self.tryWritingToSheet(data, sheetToWrite, excelRangeName);
        end
    end
    
    methods (Access = 'private')
        function newSheet = addNewSheet(self)
            numberOfSheets = self.workbookSheets.Count;
            self.workbookSheets.Add([], self.workbookSheets.Item(numberOfSheets));
            numberOfSheets = numberOfSheets + 1;
            newSheet = self.workbookSheets.Item(numberOfSheets);
            newSheet.Name = self.sheetName;
        end
    end
    
    methods (Static)
        function rangeName = getExcelRangeName(topLeftRow, topLeftCol, data)
            BottomRightCol = size(data, 2) + topLeftCol - 1;
            BottomRightRow = size(data, 1) + topLeftRow - 1;
            rangeName = ExcelSheetWriter.getRangeName(topLeftCol, BottomRightCol, topLeftRow, BottomRightRow);
        end
    end
    
    methods (Static, Access = 'private')
        function rangeName = getRangeName(firstColumn, lastColumn, firstRow, lastRow)
            firstColumnName = ExcelSheetWriter.getColumnNameFromColumnNumber(firstColumn);
            lastColumnName = ExcelSheetWriter.getColumnNameFromColumnNumber(lastColumn);
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
end

