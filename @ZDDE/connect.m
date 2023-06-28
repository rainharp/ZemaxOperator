function connect(obj, parameter)
%connect - Create connection between MATLAB and OpticStudio Zemax
%
%   Author: Tingyu Xue
%	Last Updated: 2023.6.28
%
%   This MATLAB function creates connection between MATLAB and
%   OpticStudio Zemax. Verified on OpticStudio Zemax 19.4.
%
%   PARAMETER parameter£º 
%   1. zmx file name or,
%   2. instance number of Zemax process
%
%   -------------------------Examples.m------------------------------------------------
%   ZOS = ZDDE;
%   ZOS.connect;                % Connect as extension, any instance number
%   ZOS.connect(0);             % Connect as extension, any instance number
%   ZOS.connect(1);             % Connect as extension, instance number 1
%   ZOS.connect('sample.zmx');  % Connect as standalone app and open 'sample.zmx'
%   ZOS.connect('');            % Connect as standalone app and open an new zmx file


   %% Create initial connection instance TheConnection
    import System.Reflection.*;
    import ZOSAPI.*;

    % Find current version of opticStudio
    zemaxData = winqueryreg('HKEY_CURRENT_USER', 'Software\Zemax', 'ZemaxRoot');    % get Zemax document path

    % add ZOS-API assemblies to MATLAB
    NetHelper = strcat(zemaxData, '\ZOS-API\Libraries\ZOSAPI_NetHelper.dll');       % get path of ZOSAPI_NetHelper.dll
    NET.addAssembly(NetHelper);                                                     % add ZOSAPI_NetHelper.dll to MATLAB                       
    success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize();                     % Initialize Opticstudio Zemax

    if success == 1
        disp(strcat('Found OpticStudio at: ', char(ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory())));
    else          
        error('Initialization failed, please check whether Opticstudio Zemax is installed correctly.');
        TheApplication = [];    % If failed, returns empty.
        return;
    end

    % create initial connection instance
    NET.addAssembly(AssemblyName('ZOSAPI_Interfaces'));
    NET.addAssembly(AssemblyName('ZOSAPI'));
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
            HandleError('ZDDE.m, License check failed! Please check instance id£¬or check if Zemax is already open to wait for extension connection.');
            TheApplication = [];
        end
    end

    %% Connect as standalone app
    if exist('zfile_path','var')
        Mode = 'Standalone';
        % Check if zmx file exist, if not, returns empty.
        if isempty(zfile_path)
            TheApplication = TheConnection.CreateNewApplication();
        else
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
    end
	% Set properties TheApplication and Mode.
    obj.TheApplication = TheApplication;
    obj.Mode = Mode;

end

function HandleError(message)
    warning(message);
end

