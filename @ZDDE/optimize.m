function optimize(obj)
    LocalOpt = obj.TheApplication.PrimarySystem.Tools.OpenLocalOptimization();
    LocalOpt.Algorithm = ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares;
    LocalOpt.Cycles = ZOSAPI.Tools.Optimization.OptimizationCycles.Automatic;
    LocalOpt.NumberOfCores = 8;
    LocalOpt.RunAndWaitForCompletion();
    LocalOpt.Close();
end