classdef ExcelMatlab < handle
    properties (Access = 'protected')
        app
        workbook
        sheets
        fullPathToFile
        successSaving = false
    end
    
    methods
        function self = ExcelMatlab(fullPathToFile)
            self.fullPathToFile = fullPathToFile;
            self.app = COM.Excel_Application('server', '', 'IDispatch');
            self.app.DisplayAlerts = false;
            try
                self.workbook = self.app.Workbooks.Open(fullPathToFile);
            catch
                self.workbook = self.app.Workbooks.Add();
            end
            
            try
                self.workbook.SaveAs(fullPathToFile);
                self.successSaving = true;
            catch MException
                display(MException.message);
                throw(MException);
            end
            self.sheets = self.workbook.Sheets;
        end
        
        function delete(self)
            if self.successSaving
                self.workbook.SaveAs(self.fullPathToFile);
            end
            Quit(self.app);
            delete(self.app);
        end
        
        function writeToSheet(self, data, sheet, topLeftRow, topLeftCol)
            sheetNumber = self.findSheetNumber(sheet);
            if sheetNumber
                sheetToWrite = self.sheets.Item(sheetNumber);
            else
                numberOfSheets = self.sheets.Count;
                self.sheets.Add([], self.sheets.Item(numberOfSheets));
                numberOfSheets = numberOfSheets + 1;
                sheetToWrite = self.sheets.Item(numberOfSheets);
                sheetToWrite.Name = sheet;
            end
            
            BottomRightCol = size(data, 2) + topLeftCol - 1;
            BottomRightRow = size(data, 1) + topLeftRow - 1;
            excelRange = ExcelMatlab.getExcelRangeString(topLeftCol, BottomRightCol, topLeftRow, BottomRightRow);
            rangeToWrite = get(sheetToWrite, 'Range', excelRange);
            rangeToWrite.Value = data;
        end
    end
    
    methods (Access = 'protected')
        function sheetNumber = findSheetNumber(self, sheet)
            numberOfSheets = self.sheets.Count;
            namesOfSheets = cell(1, numberOfSheets);
            for i = 1:numberOfSheets;
                namesOfSheets{i} = self.sheets.Item(i).Name;
            end

            [~, sheetNumber] = ismember(sheet, namesOfSheets);
        end
    end
    
    methods (Static)
        function rangeString = getExcelRangeString(firstColumn, lastColumn, firstRow, lastRow)
        %getExcelRangeString Returns 'RANGE' argument required for xlswrite and xlsread.
        %   rangeString = getExcelRangeString(firstColumn, lastColumn, firstRow, lastRow)
        %   returns the string corresponding to the rectangular portion of an excel
        %   spreadsheet specified by the bounds of the input arguments.
        %
        %   Examples:
        %   
        %   rangeString = getExcelRangeString(1, 5, 1, 5)
        %
        %   rangeString =
        %
        %   A1:E5
        %
        %   rangeString = getExcelRangeString(6, 9, 4, 7)
        %
        %   rangeString =
        %
        %   F4:I7
        %
        %   Note:
        %
        %   Does not support ranges that require three consecutive alphabetic
        %   letters.  Therefore the maximum value for lastColumn and lastRow is
        %   26 * 27 = 702.

        firstColumnName = getColumnNameFromColumnNumber(firstColumn);
        lastColumnName = getColumnNameFromColumnNumber(lastColumn);

        rangeString = [firstColumnName, num2str(firstRow), ':', lastColumnName, num2str(lastRow)];

            function columnName = getColumnNameFromColumnNumber(columnNumber)
                numberOfLettersInAlphabet = 26;

                if columnNumber > numberOfLettersInAlphabet
                    counter = 0;

                    while (columnNumber - numberOfLettersInAlphabet > 0)
                        columnNumber = columnNumber - numberOfLettersInAlphabet;
                        counter = counter + 1;
                    end

                    columnName = [char('A' + counter - 1), char('A' + columnNumber - 1)];
                else
                    columnName = char('A' + columnNumber - 1);
                end
            end
        end
    end
end

