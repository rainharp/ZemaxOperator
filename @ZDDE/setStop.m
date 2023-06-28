function setStop(obj, surfaceID)
    obj.TheApplication.PrimarySystem.LDE.GetSurfaceAt(surfaceID).TypeData.IsStop = 1;      % set surfaceID as stop
end
