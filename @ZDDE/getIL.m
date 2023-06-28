function IL = getIL(obj)
    % Calculate insertion loss
    % last updated: 2022.6.13
    TheSystem = obj.TheApplication.PrimarySystem;
    nsur = TheSystem.LDE.NumberOfSurfaces;
    Efficiency = TheSystem.MFE.GetOperandValue(ZOSAPI.Editors.MFE.MeritOperandType.POPD, nsur, 0, 0, 0, 0, 0, 0, 0);
    IL = -10 * log10(Efficiency);
end