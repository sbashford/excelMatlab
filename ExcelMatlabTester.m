classdef ExcelMatlabTester < matlab.unittest.TestCase   
    properties (Constant)
        TEST_FILE = 'testExcelMatlab.xlsx'
        BAD_FILES = { ...
                'word.doc', ...
                'library.dll', ...
                'library.lib', ...
                'code.m', ...
                'code.c', ...
                'image.jpg', ...
                'pdf.pdf', ...
                'hype.html' ...
                'go.exe', ...
                'jam.wav', ...
                };
    end
    
    properties (Access = 'private')
        fullPathToTestFile
        badFilesDirectory
        defaultWarning
    end
    
    methods (TestClassSetup)
        function classSetup(self)
            self.fullPathToTestFile = [pwd(), filesep, self.TEST_FILE];
            self.deleteTestFile();
            self.defaultWarning = warning('off', 'MATLAB:xlswrite:AddSheet');
            self.addTeardown(@() warning(self.defaultWarning));
            
            self.badFilesDirectory = pwd();
            for i = 1:numel(self.BAD_FILES)
                fclose(fopen([self.badFilesDirectory, filesep, self.BAD_FILES{i}], 'w'));
            end
            self.addTeardown(@self.deleteBadFiles);
        end
    end
    
    methods (TestMethodSetup)
        function methodSetup(self)
            self.addTeardown(@self.deleteTestFile);
        end
    end
    
    methods (Access = 'private')
        function deleteTestFile(self)
            if exist(self.fullPathToTestFile, 'file')
                delete(self.fullPathToTestFile);
            end
        end
        
        function deleteBadFiles(self)
            for i = 1:numel(self.BAD_FILES)
                if exist([self.badFilesDirectory, filesep, self.BAD_FILES{i}], 'file')
                    delete([self.badFilesDirectory, filesep, self.BAD_FILES{i}]);
                end
            end
        end
    end
    
    methods (Test)
        function writeNumericTest(self)
            randomArray = rand(10);
            myExcel = ExcelMatlab(self.fullPathToTestFile, 'w');
            myExcel.writeToSheet(randomArray, 'testSheet', 1, 1);
            myExcel.save();
            delete(myExcel);
            numericArray = xlsread(self.fullPathToTestFile, 'testSheet');
            self.verifyEqual(numericArray, randomArray);
        end
        
        function verifyCellPlacement(self)
            myExcel = ExcelMatlab(self.fullPathToTestFile, 'w');
            randomNumber = rand(1);
            myExcel.writeToSheet(randomNumber, 'testSheet', 17, 31);
            myExcel.save();
            delete(myExcel);
            numberRead = xlsread(self.fullPathToTestFile, 'testSheet', 'AE17:AE17');
            self.verifyEqual(numberRead, randomNumber);
        end
        
        function readNumericCell(self)
            randomNumber = rand(1);
            xlswrite(self.fullPathToTestFile, randomNumber, 'testSheet', 'B11');
            myExcel = ExcelMatlab(self.fullPathToTestFile);
            valueRead = myExcel.readCell('testSheet', 11, 2);
            delete(myExcel);
            self.verifyEqual(valueRead, randomNumber);
        end
        
        function readNumericColumn(self)
            randomCol = rand(10, 1);
            xlswrite(self.fullPathToTestFile, randomCol, 'testSheet', 'C2:C11');
            myExcel = ExcelMatlab(self.fullPathToTestFile);
            colRead = myExcel.readNumericColumnRange('testSheet', 3, 2, 11);
            delete(myExcel);
            self.verifyEqual(colRead, randomCol);
        end
        
        function assertInvalidFormat(self)
            for i = 1:numel(self.BAD_FILES)
                self.verifyBadFile(self.BAD_FILES{i});
            end
        end
        
        function assertInvalidSheetName(self)
            myExcel = ExcelMatlab(self.fullPathToTestFile, 'w');
            self.verifyError( @() writeToSheet(myExcel, [1,2], 52, 5, 5), 'ExcelMatlab:invalidSheetName');
        end
        
        function assertInvalidRowCol(self)
            myExcel = ExcelMatlab(self.fullPathToTestFile, 'w');
            self.verifyError( @() writeToSheet(myExcel, [1,2], 'testSheet', 1.1, 3), 'ExcelMatlab:invalidRowCol');
            self.verifyError( @() writeToSheet(myExcel, [1,2], 'testSheet', 0, 3), 'ExcelMatlab:invalidRowCol');
            self.verifyError( @() writeToSheet(myExcel, [1,2], 'testSheet', -1, 3), 'ExcelMatlab:invalidRowCol');
            self.verifyError( @() writeToSheet(myExcel, [1,2], 'testSheet', NaN, 3), 'ExcelMatlab:invalidRowCol');
            self.verifyError( @() writeToSheet(myExcel, [1,2], 'testSheet', Inf, 3), 'ExcelMatlab:invalidRowCol');
            self.verifyError( @() writeToSheet(myExcel, [1,2], 'testSheet', 1i, 3), 'ExcelMatlab:invalidRowCol');
        end
        
        function assertInvalidPermission(self)
            xlswrite(self.fullPathToTestFile, 1, 'testSheet');
            myExcel = ExcelMatlab(self.fullPathToTestFile);
            self.verifyError( @() writeToSheet(myExcel, [1,2], 'testSheet', 1, 1), 'ExcelMatlab:invalidPermission');
        end
        
        function assertInvalidSheet(self)
            xlswrite(self.fullPathToTestFile, 1, 'testSheet');
            myExcel = ExcelMatlab(self.fullPathToTestFile);
            self.verifyError( @() readCell(myExcel, 'notSheet', 1, 1), 'ExcelMatlab:invalidSheet');
            self.verifyError( @() readNumericColumnRange(myExcel, 'notSheet', 1, 1, 1), 'ExcelMatlab:invalidSheet');
        end
    end
    
    methods (Access = 'private')
        function verifyBadFile(self, file)
            self.verifyError( @() ExcelMatlab([self.badFilesDirectory, filesep, file], 'w'), 'ExcelMatlab:invalidFileFormat');
        end
    end
end

