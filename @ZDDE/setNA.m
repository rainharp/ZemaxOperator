function setNA(obj, NA)
    obj.TheApplication.PrimarySystem.SystemData.Aperture.ApertureType = ZOSAPI.SystemData.ZemaxApertureType.ObjectSpaceNA;
    obj.TheApplication.PrimarySystem.SystemData.Aperture.ApertureValue = NA;
    obj.TheApplication.PrimarySystem.SystemData.Aperture.ApodizationType = ZOSAPI.SystemData.ZemaxApodizationType.Gaussian;
    obj.TheApplication.PrimarySystem.SystemData.Aperture.ApodizationFactor = 1;
end