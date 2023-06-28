function Wavelength = getWavelength(obj, id)
    Wavelength = obj.TheApplication.PrimarySystem.SystemData.Wavelengths.GetWavelength(id);
end