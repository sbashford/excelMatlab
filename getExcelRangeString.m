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