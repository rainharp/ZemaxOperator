function detectorData = getDetectorData(obj, DetectorID)
    arguments
        obj
        DetectorID  (1,1) uint8 {mustBeNumeric, mustBeGreaterThan(DetectorID,0)} 
    end
    ID = DetectorID;
    TheNCE = obj.NCE;
    data = NET.createArray('System.Double', TheNCE.GetDetectorSize(ID));
    TheNCE.GetAllDetectorData(ID, 1, TheNCE.GetDetectorSize(ID), data);
    [~, rows, cols] = TheNCE.GetDetectorDimensions(ID);
    detectorData = flipud(rot90(reshape(data.double, rows, cols)));
end