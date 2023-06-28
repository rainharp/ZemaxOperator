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