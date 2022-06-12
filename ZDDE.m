classdef ZDDE
% an encapsulated MATLAB class of ZOS-API for OpticStudio Zemax 19.4 SP2
% Author：Terrence
% Update Date: 2022.6.13

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
        function obj = ZDDE(parameter)
        %   创建 MATLAB 与 Zemax 的连接
        %
        %   作者：Tingyu Xue
        %	更新日期： 2021.9.16
        %	函数版本： 1.0
        %   函数依赖： 无
        %
        %   此函数连接打开的 Zemax 进程，创建全局变量 TheApplication
        %   参数 parameter： zmx文件名或实例编号
        %   返回值 TheApplication：ZOSAPI TheApplication 对象
        %   参考：Zemax MATLAB 应用范例
        %
        %   global TheApplication;
        %   TheApplication = ZDDE();

            %% 创建初始连接对象 TheConnection
            global TheApplication;
            import System.Reflection.*;
            import ZOSAPI.*;

            % 找到当前安装的 OpticStudio 版本
            zemaxData = winqueryreg('HKEY_CURRENT_USER', 'Software\Zemax', 'ZemaxRoot');    % 获取 Zemax 文档目录
            NetHelper = strcat(zemaxData, '\ZOS-API\Libraries\ZOSAPI_NetHelper.dll');       % 获取 ZOSAPI_NetHelper.dll 路径
            % NetHelper = 'C:\Users\Documents\Zemax\ZOS-API\Libraries\ZOSAPI_NetHelper.dll';  % 自定义 ZOSAPI_NetHelper.dll 路径
            NET.addAssembly(NetHelper);                                                     % 将 ZOSAPI_NetHelper.dll 添加至 MATLAB   

            success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize();
            % success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize('C:\Program Files\OpticStudio\');  % 从自定义 Zemax 程序目录初始化
            if success == 1
                disp(strcat('Found OpticStudio at: ', char(ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory())));
            else
                % 若初始化失败，返回空
                TheApplication = [];
                return;
            end

            % 将 ZOS-API assemblies 添加至 MATLAB
            NET.addAssembly(AssemblyName('ZOSAPI_Interfaces'));
            NET.addAssembly(AssemblyName('ZOSAPI'));

            % 创建初始连接对象
            TheConnection = ZOSAPI.ZOSAPI_Connection();

            if ~exist('parameter', 'var')
                parameter = 0;
                instance = parameter;
            elseif strcmp(class(parameter), 'char')
                zfile_path = parameter;
            else
                try
                    parameter = int32(parameter);
                catch
                    parameter = 0;
                    warning('Invalid parameter {parameter}');
                end
                instance = parameter;
            end

            % 尝试创建独立连接

            % 注意 - 若提示 'Unable to load one or more of the requested types', 
            % 通常是因为尝试连接 32 位的 Zemax 和 64 位的 Matlab, 或 64 位的 
            % Matlab 和 32 位的 Zemax。这是由 MATLAB 与 .NET 的交互造成的。当前
            % 只能通过安装同为 32 位或 64 位的 Zemax 和 MATLAB 解决。

            %% 以扩展方式连接
            if exist('instance', 'var')
                Mode = 'Extension';
                TheApplication = TheConnection.ConnectAsExtension(instance);
                if isempty(TheApplication)
                   HandleError('Failed to connect to OpticStudio!');
                end
                if ~TheApplication.IsValidLicenseForAPI
                    %TheApplication.CloseApplication();
                    %HandleError('License check failed!');
                    HandleError('ZDDE.m, License check failed! 请检查实例编号，或确认是否已打开 Zemax 交互扩展等待连接。');
                    TheApplication = [];
                end
            end

            %% 以独立应用方式连接
            if exist('zfile_path','var')
                Mode = 'Standalone';
                % 判断文件是否存在，如果不存在，返回错误
                if exist(zfile_path) 
                    % 如果路径不完整，则补全路径
                    if ~strcmp(':', zfile_path(2))
                        zfile_path = fullfile(pwd, zfile_path);
                    end
                    TheApplication = TheConnection.CreateNewApplication();
                    if isempty(TheApplication)
                        ME = MXException('An unknown connection error occurred!');
                        throw(ME);
                    end
                    if ~TheApplication.IsValidLicenseForAPI
                        ME = MXException('License check failed!');
                        throw(ME);
                        TheApplication = [];
                    end

                    if isempty(TheApplication)
                        % 如果初始化连接失败
                        disp('Failed to initialize a connection!');
                    else
                        try
                            TheApplication.PrimarySystem.LoadFile(zfile_path, false);% 打开模型Zemax文件
                        catch err
                            TheApplication.CloseApplication();
                            rethrow(err);
                        end
                    end
                else
                    % 若 zfile_path 不存在
                    msgbox('+zdde\connect.m,  Zemax 文件不存在，请检查模型目录！');
                end
            end
            obj.TheApplication = TheApplication;
            obj.Mode = Mode;
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
            
        function LDE_InsertECollimator(obj, varargin)
            % 在序列模式中插入带8度面的准直器；；
            % InsertCollimator(3.850, 1.800, 'N-SF11', 0.198, 0.5);
            length = varargin{1};
            radius = varargin{2};
            material = varargin{3};
            d0 = varargin{4};
            semiDiameter = varargin{5};
            Tilt = varargin{6};
            obj.LDE.InsertNewSurfaceAt(1);
            obj.LDE.InsertNewSurfaceAt(1);
            obj.LDE.InsertNewSurfaceAt(1);

            Surface_1 = obj.LDE.GetSurfaceAt(0);
            Surface_2 = obj.LDE.GetSurfaceAt(1);
            Surface_3 = obj.LDE.GetSurfaceAt(2);
            Surface_4 = obj.LDE.GetSurfaceAt(3);

            % 设置 d0, 透镜长度
            Surface_1.Thickness = 0;
            Surface_2.Thickness = d0;
            Surface_3.Thickness = length;
            Surface_4.Thickness = 0;

            % 设置材料
            Surface_1.Material = 'F_Silica';  
            Surface_2.Material = '';  
            Surface_3.Material = material;  
            Surface_4.Material = '';

            % 设置透镜半径，曲率半径
            Surface_3.SemiDiameter = semiDiameter;
            Surface_4.SemiDiameter = semiDiameter;
            Surface_4.Radius = -radius;
            
            % 设置行颜色
            Surface_1.TypeData.RowColor = ZOSAPI.Common.ZemaxColor.Color13;
            Surface_2.TypeData.RowColor = ZOSAPI.Common.ZemaxColor.Color13;
            Surface_3.TypeData.RowColor = ZOSAPI.Common.ZemaxColor.Color13;
            Surface_4.TypeData.RowColor = ZOSAPI.Common.ZemaxColor.Color13;
            
            % 设置 8 度面
            SurfaceType_CB = Surface_2.GetSurfaceTypeSettings(ZOSAPI.Editors.LDE.SurfaceType.Tilted);
            Surface_2.ChangeType(SurfaceType_CB);
            Surface_2.SurfaceData.Y_Tangent = tan(Tilt*pi/180);
            Surface_3.ChangeType(SurfaceType_CB);
            Surface_3.SurfaceData.Y_Tangent = tan(Tilt*pi/180);

            % 设置注释
            Surface_1.Comment = '准直器';
            Surface_2.Comment = '准直器';
            Surface_3.Comment = '准直器';
            Surface_4.Comment = '准直器';
        end
        
        function LDE_InsertRCollimator(obj, varargin)
            length = varargin{1};
            radius = varargin{2};
            material = varargin{3};
            d0 = varargin{4};
            semiDiameter = varargin{5};
            Tilt = varargin{6};
            
            count = obj.LDE.NumberOfSurfaces;
            obj.LDE.InsertNewSurfaceAt(count-1);
            obj.LDE.InsertNewSurfaceAt(count-1);
            Surface_1 = obj.LDE.GetSurfaceAt(count-1);
            Surface_2 = obj.LDE.GetSurfaceAt(count);
            Surface_3 = obj.LDE.GetSurfaceAt(count+1);
            
            % 设置透镜长度，曲率，半径, d0
            Surface_1.Thickness = length;
            Surface_1.Radius = radius;
            Surface_1.SemiDiameter = semiDiameter;
            Surface_2.SemiDiameter = semiDiameter;
            Surface_3.SemiDiameter = semiDiameter;  
            Surface_2.Thickness = d0;
            
            % 设置 8 度面
            SurfaceType_CB = Surface_2.GetSurfaceTypeSettings(ZOSAPI.Editors.LDE.SurfaceType.Tilted);
            Surface_2.ChangeType(SurfaceType_CB);
            Surface_2.SurfaceData.Y_Tangent = tan(Tilt*pi/180);
            Surface_3.ChangeType(SurfaceType_CB);
            Surface_3.SurfaceData.Y_Tangent = tan(Tilt*pi/180);

            % 设置材料
            Surface_1.Material = material;  
            Surface_2.Material = '';  
            Surface_3.Material = 'F_Silica';  
            
            % 设置行颜色
            Surface_1.TypeData.RowColor = ZOSAPI.Common.ZemaxColor.Color13;
            Surface_2.TypeData.RowColor = ZOSAPI.Common.ZemaxColor.Color13;
            Surface_3.TypeData.RowColor = ZOSAPI.Common.ZemaxColor.Color13;
            
            % 设置注释
            Surface_1.Comment = '准直器';
            Surface_2.Comment = '准直器';
            Surface_3.Comment = '准直器';
        end
        
        function New(obj)
            obj.TheApplication.PrimarySystem.New(false);
        end
        
        function Open(obj, filename)
            obj.TheApplication.PrimarySystem.LoadFile(filename, false);
        end
        
        function Save(obj)
            obj.TheApplication.PrimarySystem.Save;
        end
        
        function SaveAs(obj, filename)
            obj.TheApplication.PrimarySystem.SaveAs(filename);
        end
        
        function Optimize(obj)
            LocalOpt = obj.TheApplication.PrimarySystem.Tools.OpenLocalOptimization();
            LocalOpt.Algorithm = ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares;
            LocalOpt.Cycles = ZOSAPI.Tools.Optimization.OptimizationCycles.Automatic;
            LocalOpt.NumberOfCores = 8;
            LocalOpt.RunAndWaitForCompletion();
            LocalOpt.Close();
        end
        
        function IL = getIL(obj)
            TheSystem = obj.TheApplication.PrimarySystem;
            nsur = TheSystem.LDE.NumberOfSurfaces;
            Efficiency = TheSystem.MFE.GetOperandValue(ZOSAPI.Editors.MFE.MeritOperandType.POPD, nsur, 0, 0, 0, 0, 0, 0, 0);
            IL = -10 * log10(Efficiency);
        end
        
        function setNA(obj, NA)
            obj.TheApplication.PrimarySystem.SystemData.Aperture.ApertureType = ZOSAPI.SystemData.ZemaxApertureType.ObjectSpaceNA;
            obj.TheApplication.PrimarySystem.SystemData.Aperture.ApertureValue = NA;
            obj.TheApplication.PrimarySystem.SystemData.Aperture.ApodizationType = ZOSAPI.SystemData.ZemaxApodizationType.Gaussian;
            obj.TheApplication.PrimarySystem.SystemData.Aperture.ApodizationFactor = 1;
        end
        
        function setStop(obj, surfaceID)
            obj.TheApplication.PrimarySystem.LDE.GetSurfaceAt(surfaceID).TypeData.IsStop = 1;      % 设置面 surfaceID 为光阑
        end
        
        function setWavelength(obj, wavelength)
            obj.TheApplication.PrimarySystem.SystemData.Wavelengths.GetWavelength(1).Wavelength = wavelength;
            obj.TheApplication.PrimarySystem.SystemData.Wavelengths.GetWavelength(1).MakePrimary;
        end
        
        function varargout = getMFE(obj,varargin)
            %getMFE - 获取 OpticStudio Zemax 的评价函数表（值）
            %
            %	作者：Tingyu Xue
            %	更新日期： 2021.7.24
            %	函数版本： 1.0
            %     函数依赖： 无
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
                    disp('未找到 ZOSAPI_Application 对象！');
                    return;
                elseif strcmp(class(TheApplication),'ZemaxUI.Common.ViewModels.ZOSAPI_Application')
                    N = TheApplication.PrimarySystem.MFE.NumberOfOperands;    % 评价函数数量
                    TheMFE = TheApplication.PrimarySystem.MFE;
                    TheMFE.CalculateMeritFunction();                          % 计算评价函数
                    if nargin < 2          
                        for ii = 1: N
                            Type{ii,1} = char(TheApplication.PrimarySystem.MFE.GetOperandAt(ii).RowTypeName);
                            Target(ii,1) = TheApplication.PrimarySystem.MFE.GetOperandAt(ii).Target;
                            Weight(ii,1) = TheApplication.PrimarySystem.MFE.GetOperandAt(ii).Weight;
                            Value(ii,1) = TheApplication.PrimarySystem.MFE.GetOperandAt(ii).Value;
                        end
                        MFETable = table(Type,Target,Weight, Value);   % 返回评价函数表
                        varargout{1} = MFETable;
                    elseif nargin == 2
                        id = varargin{1};
                        if length(id) == 1
                            if id==fix(id) && id <= N && id > 0
                                varargout{1} = TheMFE.GetOperandAt(id).Value; % 返回第id行评价函数的值
                            else
                                disp('评价函数序号错误！');
                                return;
                            end
                        else
                            for i = 1:length(id)
                                if id(i)==fix(id(i)) && id(i) <= N && id(i) > 0
                                    varargout{i} = TheMFE.GetOperandAt(id(i)).Value;  % 返回第id(i)行评价函数的值
                                else
                                    disp('评价函数序号错误！');
                                    return;
                                end
                            end
                        end
                    end
                else
                    disp('变量类型错误。全局变量 TheApplication 中应保存 ZOSAPI_Application 对象! ');
                    return;
                end
        end
        
        function Surface = getSurface(obj,id)
            Surface = obj.LDE.GetSurfaceAt(id);
        end
        
        function NSCTrace(obj, FileName)
            % 非序列光线追迹
            File = obj.TheSystem.SystemFile;               % Zemax 文件路径
            obj.TheSystem.LoadFile(File, false);
            if ~isempty(obj.TheSystem.Tools.CurrentTool)
                obj.TheSystem.Tools.CurrentTool.Close();
            end
            NSCRayTrace = obj.TheSystem.Tools.OpenNSCRayTrace();
            NSCRayTrace.SplitNSCRays = true;
            NSCRayTrace.ScatterNSCRays = false;
            NSCRayTrace.UsePolarization = true;
            NSCRayTrace.IgnoreErrors = true;
            NSCRayTrace.SaveRays = true;
            NSCRayTrace.SaveRaysFile = FileName;
            NSCRayTrace.ClearDetectors(0);
            NSCRayTrace.RunAndWaitForCompletion();
            NSCRayTrace.Close();
        end
        
        function detectorData = getDetectorData(obj, DetectorID)
            ID = DetectorID;
            TheNCE = obj.NCE;
            data = NET.createArray('System.Double', TheNCE.GetDetectorSize(ID));
            TheNCE.GetAllDetectorData(ID, 1, TheNCE.GetDetectorSize(ID), data);
            [~, rows, cols] = TheNCE.GetDetectorDimensions(ID);
            detectorData = flipud(rot90(reshape(data.double, rows, cols)));
        end
        
        function plotDetector(obj, DetectorID)
            detectorData = obj.getDetectorData(DetectorID);
            figure('menubar', 'none', 'numbertitle', 'off');
            mesh(detectorData);
            xlabel('Column #');
            ylabel('Row #');
            view(0,90);
            axis equal;
            colorbar;
        end
        
        function Result = getZRDResult(obj, ZRDFilePath)
        % 读取 ZRD 文件
            if ~strcmp(ZRDFilePath(2),':')
                ZRDFilePath = fullfile(pwd, ZRDFilePath);
            end
            TheSystem = obj.TheSystem;
            if ~isempty(TheSystem.Tools.CurrentTool)
                TheSystem.Tools.CurrentTool.Close();
            end
            ZRDReader = TheSystem.Tools.OpenRayDatabaseReader();
            ZRDReader.ZRDFile = ZRDFilePath;
            ZRDReader.RunAndWaitForCompletion();
            if ZRDReader.Succeeded == 0
                disp('ZRD 文件读取失败!');
                disp(ZRDReader.ErrorMessage);
            else
                disp('ZRD 文件读取成功!');
            end
            Result = ZRDReader.GetResults();
        end
        
        function makeVariable(obj,surfaceID, paraName)
            surface = obj.LDE.GetSurfaceAt(surfaceID);
            eval(['surface.SurfaceData.',paraName,'_Cell.MakeSolveVariable();']);
        end
        
        function makeFixed(obj, surfaceID, paraName)
            surface = obj.LDE.GetSurfaceAt(surfaceID);
            eval(['surface.SurfaceData.',paraName,'_Cell.MakeSolveFixed();']);
        end
    end
end