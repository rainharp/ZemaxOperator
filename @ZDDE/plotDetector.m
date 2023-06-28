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