function CoatingString = getCoatingString(obj, T)
%getCoatingString - generate coating string from Transmittance
% Tingyu Xue
% Last updated: 2022.6.28
    arguments
        obj   
        T     (1,1)  {mustBeNumeric, mustBeGreaterThanOrEqual(T,0), mustBeLessThanOrEqual(T,1)} 
    end
    CoatingString = num2str(T, '%.3f');
    CoatingString = ['T',CoatingString(2:end)];
end