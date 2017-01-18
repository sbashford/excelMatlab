classdef ExcelMatlabTester < matlab.unittest.TestCase   
    properties (Constant)
        TEST_FILE = 'testExcelMatlab.xlsx'
    end
    
    properties (Access = 'private')
        fullPathToTestFile
    end
    
    methods (TestClassSetup)
        function somethingElseOrWhatever(self)
            self.fullPathToTestFile = [pwd(), '\', self.TEST_FILE];
            self.deleteTestFile();
            self.addTeardown(@self.deleteTestFile);
        end
    end
    
    methods (Access = 'private')
        function deleteTestFile(self)
            if exist(self.fullPathToTestFile, 'file')
                delete(self.fullPathToTestFile);
            end
        end
    end
    
    methods (Test)
        function writeNumericTest(self)
            randomArray = rand(10);
            myExcel = ExcelMatlab(self.fullPathToTestFile);
            myExcel.writeToSheet(randomArray, 'testSheet', 1, 1);
            delete(myExcel);
            numericArray = xlsread(self.fullPathToTestFile, 'testSheet');
            self.verifyEqual(randomArray, numericArray);
        end
    end
end

