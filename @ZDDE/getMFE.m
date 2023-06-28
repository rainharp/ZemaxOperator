function varargout = getMFE(obj,varargin)
    %getMFE - get merit function table (value)
    %
    %	Author: Tingyu Xue
    %	Last updated: 2021.7.24
    %	Version: 1.0
    %
    %	此 MATLAB 函数 在已建立储存于 TheApplication 全局变量的 ZOSAPI_Application
    %   对象的基础上，计算并返回 Zemax 评价函数表（值）。无输入参数时，计算并返回评
    %   价函数表。有 1 个输入参数时，返回某行的评价函数值，有多个输入参数时，返回各
    %   行评价函数。
    %
    %   MFETable = getMFE();
    %   MFEValue = getMFE(rowNum);
    %   [MFEValue1, MFEValue2] = getMFE(rowNum1, rowNum2);
    %   ...

        TheApplication = obj.TheApplication;
        if isempty(TheApplication)
            disp('ZOSAPI_Application object not find.');
            return;
        elseif strcmp(class(TheApplication),'ZemaxUI.Common.ViewModels.ZOSAPI_Application')
            N = TheApplication.PrimarySystem.MFE.NumberOfOperands;    % operand count
            TheMFE = TheApplication.PrimarySystem.MFE;
            TheMFE.CalculateMeritFunction();                          % calculate merit function
            if nargin < 2          
                for ii = 1: N
                    Type{ii,1} = char(TheApplication.PrimarySystem.MFE.GetOperandAt(ii).RowTypeName);
                    Target(ii,1) = TheApplication.PrimarySystem.MFE.GetOperandAt(ii).Target;
                    Weight(ii,1) = TheApplication.PrimarySystem.MFE.GetOperandAt(ii).Weight;
                    Value(ii,1) = TheApplication.PrimarySystem.MFE.GetOperandAt(ii).Value;
                end
                MFETable = table(Type,Target,Weight, Value);   % returns merit function table
                varargout{1} = MFETable;
            elseif nargin == 2
                id = varargin{1};
                if length(id) == 1
                    if id==fix(id) && id <= N && id > 0
                        varargout{1} = TheMFE.GetOperandAt(id).Value;         % return MFE value at row id
                    else
                        disp('MFE ID error！');
                        return;
                    end
                else
                    for i = 1:length(id)
                        if id(i)==fix(id(i)) && id(i) <= N && id(i) > 0
                            varargout{i} = TheMFE.GetOperandAt(id(i)).Value;  % return MFE value at row id(i)
                        else
                            disp('MFE ID error！');
                            return;
                        end
                    end
                end
            end
        else
            disp('Variable type error. The ZOSAPI_Application object should be save in global variable TheApplication.');
            return;
        end
end