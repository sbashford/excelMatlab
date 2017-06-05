classdef ExcelCOM < handle
    properties (Access = private)
        app
    end
    
    properties
        Workbooks
    end
    
    methods
        function self = ExcelCOM()
            self.app = COM.Excel_Application('server', '', 'IDispatch');
            self.app.DisplayAlerts = false;
            self.Workbooks = self.app.Workbooks();
        end
        
        function delete(self)
            self.app.Quit();
            delete(self.app);
        end
    end
end