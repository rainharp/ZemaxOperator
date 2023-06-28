function setWavelength(obj, wavelength)
    %setWavelength - set Primary Wavelength
    % last updated: 2022.6.13
    obj.TheApplication.PrimarySystem.SystemData.Wavelengths.GetWavelength(1).Wavelength = wavelength;
    obj.TheApplication.PrimarySystem.SystemData.Wavelengths.GetWavelength(1).MakePrimary;
end