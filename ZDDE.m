classdef ZDDE < handle
% ZDDE - An encapsulated MATLAB class of ZOS-API for OpticStudio Zemax 19.4 SP2
% Author: Terrence Xue
% Last Updated: 2022.6.13

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
        %   Create connection between MATLAB and OpticStudio Zemax
        %
        %   Author: Terrence
        %	Last Updated: 2021.9.16
        %	Version: 1.0
        %
        %   This MATLAB function connects to openning Zemax process by
        %   creating an instance of class ZDDE.
        %
        %   PARAMETER parameter： zmx file name or instance number of Zemax process
        %   RETURN TheApplication：ZOSAPI TheApplication Object
        %
        %   ZOS = ZDDE();

           %% Create initial connection instance TheConnection
            global TheApplication;
            import System.Reflection.*;
            import ZOSAPI.*;

            % Find current version of opticStudio
            zemaxData = winqueryreg('HKEY_CURRENT_USER', 'Software\Zemax', 'ZemaxRoot');    % get Zemax document path
            NetHelper = strcat(zemaxData, '\ZOS-API\Libraries\ZOSAPI_NetHelper.dll');       % get path of ZOSAPI_NetHelper.dll
            NET.addAssembly(NetHelper);                                                     % add ZOSAPI_NetHelper.dll to MATLAB 
            success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize();                     % Initialize Opticstudio Zemax
         
            if success == 1
                disp(strcat('Found OpticStudio at: ', char(ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory())));
            else
                % If faild, returns empty.
                TheApplication = [];
                return;
            end

            % add ZOS-API assemblies to MATLAB
            NET.addAssembly(AssemblyName('ZOSAPI_Interfaces'));
            NET.addAssembly(AssemblyName('ZOSAPI'));

            % create initial connection instance
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

            % Try create standalone connection

            % NOTE - if this fails with a message like 'Unable to load one or more of
            % the requested types', it is usually caused by try to connect to a 32-bit
            % version of OpticStudio from a 64-bit version of MATLAB (or vice-versa).
            % This is an issue with how MATLAB interfaces with .NET, and the only
            % current workaround is to use 32- or 64-bit versions of both applications.

            %% Connect as extension
            if exist('instance', 'var')
                Mode = 'Extension';
                TheApplication = TheConnection.ConnectAsExtension(instance);
                if isempty(TheApplication)
                   HandleError('Failed to connect to OpticStudio!');
                end
                if ~TheApplication.IsValidLicenseForAPI
                    %TheApplication.CloseApplication();
                    %HandleError('License check failed!');
                    HandleError('ZDDE.m, License check failed! Please check instance id，or check if Zemax is already open to wait for extension connection.');
                    TheApplication = [];
                end
            end

            %% Connect as standalone app
            if exist('zfile_path','var')
                Mode = 'Standalone';
                % Check if zmx file exist, if not, returns empty.
                if exist(zfile_path) 
                    % Complete the path if it's incomplete.
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
                        % If failed
                        disp('Failed to initialize a connection!');
                    else
                        try
                            TheApplication.PrimarySystem.LoadFile(zfile_path, false); % Load zmx file.
                        catch err
                            TheApplication.CloseApplication();
                            rethrow(err);
                        end
                    end
                else
                    % if zfile_path do not exist
                    msgbox('ZDDE.m,  zmx file do not exist, please check the zmx file path.');
                end
            end
            
            % Set properties TheApplication and Mode.
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
            % Insert collimator with tilt surface in sequence mode.
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

            % Set d0 and length of c-lens
            Surface_1.Thickness = 0;
            Surface_2.Thickness = d0;
            Surface_3.Thickness = length;
            Surface_4.Thickness = 0;

            % Set material
            Surface_1.Material = 'F_Silica';  
            Surface_2.Material = '';  
            Surface_3.Material = material;  
            Surface_4.Material = '';

            % Set semidiameter and radius of c-lens
            Surface_3.SemiDiameter = semiDiameter;
            Surface_4.SemiDiameter = semiDiameter;
            Surface_4.Radius = -radius;
            
            % Set Row Color
            Surface_1.TypeData.RowColor = ZOSAPI.Common.ZemaxColor.Color13;
            Surface_2.TypeData.RowColor = ZOSAPI.Common.ZemaxColor.Color13;
            Surface_3.TypeData.RowColor = ZOSAPI.Common.ZemaxColor.Color13;
            Surface_4.TypeData.RowColor = ZOSAPI.Common.ZemaxColor.Color13;
            
            % Set tilt surface (8deg)
            SurfaceType_CB = Surface_2.GetSurfaceTypeSettings(ZOSAPI.Editors.LDE.SurfaceType.Tilted);
            Surface_2.ChangeType(SurfaceType_CB);
            Surface_2.SurfaceData.Y_Tangent = tan(Tilt*pi/180);
            Surface_3.ChangeType(SurfaceType_CB);
            Surface_3.SurfaceData.Y_Tangent = tan(Tilt*pi/180);

            % Set Comment
            Surface_1.Comment = 'Collimator';
            Surface_2.Comment = 'Collimator';
            Surface_3.Comment = 'Collimator';
            Surface_4.Comment = 'Collimator';
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
            
            % Set lens length, radius, semidiameter and d0
            Surface_1.Thickness = length;
            Surface_1.Radius = radius;
            Surface_1.SemiDiameter = semiDiameter;
            Surface_2.SemiDiameter = semiDiameter;
            Surface_3.SemiDiameter = semiDiameter;  
            Surface_2.Thickness = d0;
            
            % Set tilt surface (8deg)
            SurfaceType_CB = Surface_2.GetSurfaceTypeSettings(ZOSAPI.Editors.LDE.SurfaceType.Tilted);
            Surface_2.ChangeType(SurfaceType_CB);
            Surface_2.SurfaceData.Y_Tangent = tan(Tilt*pi/180);
            Surface_3.ChangeType(SurfaceType_CB);
            Surface_3.SurfaceData.Y_Tangent = tan(Tilt*pi/180);

            % Set material
            Surface_1.Material = material;  
            Surface_2.Material = '';  
            Surface_3.Material = 'F_Silica';  
            
            % Set row color
            Surface_1.TypeData.RowColor = ZOSAPI.Common.ZemaxColor.Color13;
            Surface_2.TypeData.RowColor = ZOSAPI.Common.ZemaxColor.Color13;
            Surface_3.TypeData.RowColor = ZOSAPI.Common.ZemaxColor.Color13;
            
            % Set comment
            Surface_1.Comment = 'Collimator';
            Surface_2.Comment = 'Collimator';
            Surface_3.Comment = 'Collimator';
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
            % Calculate insertion loss
            % last updated: 2022.6.13
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
            obj.TheApplication.PrimarySystem.LDE.GetSurfaceAt(surfaceID).TypeData.IsStop = 1;      % set surfaceID as stop
        end
        
        function setWavelength(obj, wavelength)
            %setWavelength - set Primary Wavelength
            % last updated: 2022.6.13
            obj.TheApplication.PrimarySystem.SystemData.Wavelengths.GetWavelength(1).Wavelength = wavelength;
            obj.TheApplication.PrimarySystem.SystemData.Wavelengths.GetWavelength(1).MakePrimary;
        end
        
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
        
        function Surface = getSurface(obj,id)
            Surface = obj.LDE.GetSurfaceAt(id);
        end
        
        function NSCTrace(obj, FileName)
            % NSC Ray Tracing
            File = obj.TheSystem.SystemFile;              % Zemax File Path
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
        % Read ZRD Files
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
                disp('ZRD File Reading Failed!');
                disp(ZRDReader.ErrorMessage);
            else
                disp('ZRD File Reading Succeed!');
            end
            Result = ZRDReader.GetResults();
        end
        
        function RayTable = getRayTable(obj, ZRDFilePath, varargin)
        %getRayTable - Read ZRD Files and save ray data to RayTable
        % last updated: 2022.6.13
            if nargin == 2
                filterString = "";
                option = false;                
            elseif nargin == 3
                filterString = varargin{1};
                option = false;
            else
                filterString = varargin{1};
                option = varargin{2};
            end
            if ~strcmp(ZRDFilePath(2),':')
                ZRDFilePath = fullfile(pwd, ZRDFilePath);
            end
            TheSystem = obj.TheSystem;
            if ~isempty(TheSystem.Tools.CurrentTool)
                TheSystem.Tools.CurrentTool.Close();
            end
            ZRDReader = TheSystem.Tools.OpenRayDatabaseReader();
            ZRDReader.ZRDFile = ZRDFilePath;
            ZRDReader.Filter = filterString;
            ZRDReader.RunAndWaitForCompletion();
            if ZRDReader.Succeeded == 0
                disp('ZRD File Reading Failed!');
                disp(ZRDReader.ErrorMessage);
            else
                disp('ZRD File Reading Succeed!');
            end
            ZRDResult = ZRDReader.GetResults();
            [success_NextResult, rayNumber, waveIndex, wlUM, numSegments] = ZRDResult.ReadNextResult();
            BeamCounter = 1;                                             % Light beam counter
            RowCounter = 1;                                              % Row counter

            while success_NextResult == 1
                SegmentCounter = 1;                                      % Segment counter
                [t_success_NextSegmentFull, t_segmentLevel, t_segmentParent, t_hitObj, t_hitFace, t_insideOf, t_status, ...
                      t_x, t_y, t_z, t_l, t_m, t_n, t_exr, t_exi, t_eyr, t_eyi, t_ezr, t_ezi, t_intensity, t_pathLength,...
                      t_xybin, t_lmbin, t_xNorm, t_yNorm, t_zNorm, t_index, t_startingPhase, t_phaseOf, t_phaseAt] = ZRDResult.ReadNextSegmentFull();   

                while t_success_NextSegmentFull == 1
                    [t_success_NextSegmentFull, t_segmentLevel, t_segmentParent, t_hitObj, t_hitFace, t_insideOf, t_status,...
                    t_x, t_y, t_z, t_l, t_m, t_n, t_exr, t_exi, t_eyr, t_eyi, t_ezr, t_ezi, t_intensity, t_pathLength,...
                    t_xybin, t_lmbin, t_xNorm, t_yNorm, t_zNorm, t_index, t_startingPhase, t_phaseOf, t_phaseAt] = ZRDResult.ReadNextSegmentFull(); 
                    if t_success_NextSegmentFull == 1
                        Beam(RowCounter,1) = BeamCounter;
                        Segment(RowCounter,1) = SegmentCounter;
                        SegmentLevel(RowCounter,1) = t_segmentLevel;
                        SegmentParent(RowCounter,1) = t_segmentParent;
                        hitObj(RowCounter,1) = t_hitObj;
                        hitFace(RowCounter,1) = t_hitFace;
                        insideOf(RowCounter,1) = t_insideOf;
                        status{RowCounter,1} = string(t_status);
                        x(RowCounter,1) = t_x;
                        y(RowCounter,1) = t_y;
                        z(RowCounter,1) = t_z;
                        l(RowCounter,1) = t_l;     
                        m(RowCounter,1) = t_m;     
                        n(RowCounter,1) = t_n;  
                        exr(RowCounter,1) = t_exr;
                        exi(RowCounter,1) = t_exi;
                        eyr(RowCounter,1) = t_eyr;
                        eyi(RowCounter,1) = t_eyi;
                        ezr(RowCounter,1) = t_ezr;
                        ezi(RowCounter,1) = t_ezi;
                        intensity(RowCounter,1) = t_intensity;
                        pathLength(RowCounter,1) = t_pathLength;
                        xybin(RowCounter,1) = t_xybin;
                        lmbin(RowCounter,1) = t_lmbin;
                        xNorm(RowCounter,1) = t_xNorm;
                        yNorm(RowCounter,1) = t_yNorm;
                        zNorm(RowCounter,1) = t_zNorm;
                        index(RowCounter,1) = t_index;
                        startingPhase(RowCounter,1) = t_startingPhase;
                        phaseOf(RowCounter,1) = t_phaseOf;
                        phaseAt(RowCounter,1) = t_phaseAt;

                        SegmentCounter = SegmentCounter + 1;
                        RowCounter = RowCounter + 1;    
                    end            
                end

                % Read Next Beam
                [success_NextResult, rayNumber, waveIndex, wlUM, numSegments] = ZRDResult.ReadNextResult();
                BeamCounter = BeamCounter + 1;
            end

            RayTable = table(Beam, Segment, SegmentLevel, SegmentParent, hitObj, hitFace, insideOf, status,...
                         x, y, z, l, m, n, exr, exi, eyr, eyi, ezr, ezi, intensity, pathLength,...
                         xybin, lmbin, xNorm, yNorm, zNorm, index, startingPhase, phaseOf, phaseAt);

            if option
                if exist('RayData.xlsx')
                    delete('RayData.xlsx');
                end
                writetable(RayTable, 'RayData.xlsx');
            end
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