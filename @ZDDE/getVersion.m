function Version = getVersion(obj)
%GetVersion - get the version of Opticstudio Zemax
    Version = obj.TheSystem.Server.OpticStudioVersion;
end
