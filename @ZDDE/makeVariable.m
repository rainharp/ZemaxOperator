function makeVariable(obj,surfaceID, paraName)
    surface = obj.LDE.GetSurfaceAt(surfaceID);
    eval(['surface.SurfaceData.',paraName,'_Cell.MakeSolveVariable();']);
end