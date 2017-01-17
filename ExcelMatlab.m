classdef ExcelMatlab < handle
    properties (Access = 'protected')
        excelApplication
        excelFileWorkbook
        excelFileSheets
        fullPathToFile
    end
    
    methods
        function self = ExcelMatlab(file)
            stack = dbstack('-completenames', 1);
            if size(stack)
                self.fullPathToFile = [stack(1).file, '\', file];
            else
                self.fullPathToFile = [pwd(), '\', file];
            end
            
            self.excelApplication = actxserver('Excel.Application');
            self.excelApplication.DisplayAlerts = false;
            try
                self.excelFileWorkbook = self.excelApplication.Workbooks.Open(self.fullPathToFile);
            catch
                self.excelFileWorkbook = self.excelApplication.Workbooks.Add();
            end
            self.excelFileSheets = self.excelFileWorkbook.Sheets;
        end
        
        function delete(self)
            try
                self.excelFileWorkbook.SaveAs(self.fullPathToFile);
            catch
                fprintf('Excel file did not save successfully\n');
            end
            Quit(self.excelApplication);
            delete(self.excelApplication);
        end
        
        function writeToSheet(self, data, sheet, topLeftRow, topLeftCol)
            sheetNumber = self.findSheetNumber(sheet);
            if sheetNumber
                sheetToWrite = self.excelFileSheets.Item(sheetNumber);
            else
                numberOfSheets = self.excelFileSheets.Count;
                self.excelFileSheets.Add([], self.excelFileSheets.Item(numberOfSheets));
                numberOfSheets = numberOfSheets + 1;
                sheetToWrite = self.excelFileSheets.Item(numberOfSheets);
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
            numberOfSheets = self.excelFileSheets.Count;
            namesOfSheets = cell(1, numberOfSheets);
            for i = 1:numberOfSheets;
                namesOfSheets{i} = self.excelFileSheets.Item(i).Name;
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

