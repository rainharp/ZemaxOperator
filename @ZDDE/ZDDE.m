classdef ZDDE < handle
% ZDDE - An encapsulated MATLAB class of ZOS-API for OpticStudio Zemax 19.4 SP2
%        and MATLAB version over 2019b.
%
% Author: Tingyu Xue
% Last Updated: 2023.6.28

    properties
        Mode                 
        TheApplication       
    end
    
    properties(Dependent)
        TheSystem
        LDE
        NCE
        MFE
        MCE
    end
    
    methods
        function obj = ZDDE()

        end
        
        function TheSystem = get.TheSystem(obj)
            TheSystem = obj.TheApplication.PrimarySystem;
        end
        
        function LDE = get.LDE(obj)
            LDE = obj.TheApplication.PrimarySystem.LDE;
        end
        
        function MFE = get.MFE(obj)
            MFE = obj.TheApplication.PrimarySystem.MFE;
        end
        
        function NCE = get.NCE(obj)
            NCE = obj.TheApplication.PrimarySystem.NCE;
        end
        
        function MCE = get.MCE(obj)
            MCE = obj.TheApplication.PrimarySystem.MCE;
        end

    end
end

% If you shed tears when you miss the sun, you also miss the stars.