function insertECollimator_LDE(obj, varargin)
    % Insert collimator with tilt surface in sequence mode.
    % InsertECollimator(3.850, 1.800, 'N-SF11', 0.198, 0.5, 8);
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