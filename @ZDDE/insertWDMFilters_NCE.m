function insertWDMFilters_NCE(obj, StartRow, FilterThickness, FilterWidth, FilterHeight, ...
                              ChannelCount, IncidentAngle, Pitch)
   %% InsertWDMFilters.m
    %
    % This function inserts WDM Filters in non-sequential mode of OpticStudio
    % Zemax.
    %
    % Author: Terrence
    % Last updated: 2022.6.16
    %
%             StartRow = 1;
%             FilterThickness = 1.0;         % Thickness of Filter, Unit: mm
%             FilterWidth = 1.4;             % Width of Filter, Unit: mm
%             FilterHeight = 1.4;            % Height of Filter, Unit: mm
%             ChannelCount = 8;              % Channel Count
%             IncidentAngle = 5.2;           % Incident Angle, Unit: deg
%             Pitch = 2.0;                   % Beam Pitch, Unit: mm

    FilterXPitch = Pitch / cos(deg2rad(IncidentAngle));              % Filter X Pitch, Unit: mm
    FilterZPitch = -FilterXPitch / 2 / tan(deg2rad(IncidentAngle));  % Filter Z Pitch, Unit: mm

    TheNCE = obj.NCE;
    NumberOfObject = obj.NCE.NumberOfObjects;

    % Insert Ref Object
    TheNCE.InsertNewObjectAt(StartRow);                                                             % Insert New Object      
    RefObj = TheNCE.GetObjectAt(StartRow);                                                          % Get Object
    ObjectType_SF = RefObj.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.Sphere);             % Get Sphere Object Type
    RefObj.ChangeType(ObjectType_SF);                                                               % Change Object Type to Sphere Volume
    RefObj.ObjectData.Radius = 0.10;                                                                % Set Radius
    RefObj.TypeData.RowColor = ZOSAPI.Common.ZemaxColor.Color13;                                    % Set Row Color
    RefObj.Comment = 'WDM Filter Ref Object';                                                       % Set Comment
    RefObj.DrawData.DoNotDrawObject = 0;                                                            % Do not draw

    % Insert WDM Filters
    TheNCE.InsertNewObjectAt(StartRow+1);                                                           % Insert New Object      
    Filter = TheNCE.GetObjectAt(StartRow+1);                                                        % Get Object
    ObjectType_RT = Filter.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.RectangularVolume);  % Get RectangularVolume Object Type
    Filter.ChangeType(ObjectType_RT);                                                               % Change Object Type to Rectangular Volume
    Filter.ObjectData.X1HalfWidth = FilterWidth/2;                                                  % Set Filter Width
    Filter.ObjectData.X2HalfWidth = FilterWidth/2;                                                  % Set Filter Width
    Filter.ObjectData.Y1HalfWidth = FilterHeight/2;                                                 % Set Filter Height
    Filter.ObjectData.Y2HalfWidth = FilterHeight/2;                                                 % Set Filter Height
    Filter.ObjectData.ZLength = FilterThickness;                                                    % Set Filter Thickness
    Filter.TypeData.RowColor = ZOSAPI.Common.ZemaxColor.Color13;                                    % Set Row Color
    Filter.Comment = 'WDM Filter';                                                                  % Set Comment
    Filter.Material = 'WMS-15';                                                                     % Set Filter Material
    Filter.RefObject = StartRow;                                                                    % Set Ref Object

    for ii = 2: ChannelCount
        obj.NCE.CopyObjects(StartRow+1, 1, ii+StartRow);
    end

    for ii = 1: ChannelCount
        objectID = StartRow + ii;
        Filter = TheNCE.GetObjectAt(objectID); 
        if mod(ii, 2) == 0
            Filter.ZPosition = FilterZPitch / 2 - FilterThickness/2;
        else
            Filter.ZPosition = - FilterZPitch / 2 - FilterThickness/2;
        end
        Filter.XPosition = FilterXPitch/2 * (ii-1) - FilterXPitch/2 * (ChannelCount-1)/2;
    end
end