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
            for i = 1:numberOfSheets
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
            excelRange = ExcelSheetWriter.getExcelRange(topLeftRow, topLeftCol, data);
            self.tryWritingToSheet(data, sheetToWrite, excelRange);
        end
    end
    
    methods (Access = 'private')
        function sheetToWrite = addNewSheet(self)
            numberOfSheets = self.workbookSheets.Count;
            self.workbookSheets.Add([], self.workbookSheets.Item(numberOfSheets));
            numberOfSheets = numberOfSheets + 1;
            sheetToWrite = self.workbookSheets.Item(numberOfSheets);
            sheetToWrite.Name = self.sheetName;
        end
        
        function tryWritingToSheet(~, data, sheetToWrite, excelRange)
            try
                rangeToWrite = get(sheetToWrite, 'Range', excelRange);
                rangeToWrite.Value = data;
            catch MException
                fprintf('Invalid range specified\n');
                throw(MException);
            end
        end
    end
    
    methods (Static)
        function rangeString = getExcelRange(topLeftRow, topLeftCol, data)
            BottomRightCol = size(data, 2) + topLeftCol - 1;
            BottomRightRow = size(data, 1) + topLeftRow - 1;
            rangeString = ExcelSheetWriter.getExcelRangeString(topLeftCol, BottomRightCol, topLeftRow, BottomRightRow);
        end
    end
    
    methods (Static, Access = 'private')
        function rangeString = getExcelRangeString(firstColumn, lastColumn, firstRow, lastRow)
            firstColumnString = ExcelSheetWriter.getColumnStringFromColumnNumber(firstColumn);
            lastColumnString = ExcelSheetWriter.getColumnStringFromColumnNumber(lastColumn);
            rangeString = [firstColumnString, num2str(firstRow), ':', lastColumnString, num2str(lastRow)];
        end
        
        function columnString = getColumnStringFromColumnNumber(columnNumber)
            numberOfLettersInAlphabet = 26;
            if columnNumber > numberOfLettersInAlphabet
                counter = 0;
                
                while (columnNumber - numberOfLettersInAlphabet > 0)
                    columnNumber = columnNumber - numberOfLettersInAlphabet;
                    counter = counter + 1;
                end
                
                columnString = [char('A' + counter - 1), char('A' + columnNumber - 1)];
            else
                columnString = char('A' + columnNumber - 1);
            end
        end
    end
end

