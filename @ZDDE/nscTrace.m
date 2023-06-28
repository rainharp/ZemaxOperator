function nscTrace(obj, FileName)
%nscTrace - Perform NSC Tracing and save ZRD file.
% Author: Tingyu Xue
% Last updated: 2022.6.15

    % Close opening Tool
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