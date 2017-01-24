classdef ExcelMatlab < handle
    properties (Access = 'private')
        app
        workbook
        fullPathToFile
        successSaving = false
    end
    
    methods
        function self = ExcelMatlab(fullPathToFile)
            assert(ischar(fullPathToFile), 'Path must be a string');
            self.fullPathToFile = fullPathToFile;
            self.startExcel();
            self.openWorkbook();
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
            workbookSheets = self.workbook.Sheets;
            sheetWriter = ExcelSheetWriter(workbookSheets, sheetName);
            sheetWriter.write(data, topLeftRow, topLeftCol);
        end
        
        function delete(self)
            if self.successSaving
                self.workbook.SaveAs(self.fullPathToFile);
            end
            Quit(self.app);
            delete(self.app);
        end
    end
end

