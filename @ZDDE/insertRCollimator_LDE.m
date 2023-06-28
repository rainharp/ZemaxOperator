function insertRCollimator_LDE(obj, varargin)
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