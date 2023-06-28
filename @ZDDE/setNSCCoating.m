function setNSCCoating(obj, objectID, surfaceID, coatingString)
%setCoating - setting coating of NSC object
% Author: Tingyu Xue
% Last updated: 2022.6.15
    obj.NCE.GetObjectAt(objectID).CoatScatterData.GetFaceData(surfaceID).Coating = coatingString;
end