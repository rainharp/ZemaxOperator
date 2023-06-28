function RayTable = getRayTable(obj, ZRDFilePath, varargin)
%getRayTable - Read ZRD Files and save ray data to RayTable
% last updated: 2022.6.13
    if nargin == 2
        filterString = "";
        option = false;                
    elseif nargin == 3
        filterString = varargin{1};
        option = false;
    else
        filterString = varargin{1};
        option = varargin{2};
    end
    if ~strcmp(ZRDFilePath(2),':')
        ZRDFilePath = fullfile(pwd, ZRDFilePath);
    end
    TheSystem = obj.TheSystem;
    if ~isempty(TheSystem.Tools.CurrentTool)
        TheSystem.Tools.CurrentTool.Close();
    end
    ZRDReader = TheSystem.Tools.OpenRayDatabaseReader();
    ZRDReader.ZRDFile = ZRDFilePath;
    ZRDReader.Filter = filterString;
    ZRDReader.RunAndWaitForCompletion();
    if ZRDReader.Succeeded == 0
        disp('ZRD File Reading Failed!');
        disp(ZRDReader.ErrorMessage);
    else
        disp('ZRD File Reading Succeed!');
    end
    ZRDResult = ZRDReader.GetResults();
    [success_NextResult, rayNumber, waveIndex, wlUM, numSegments] = ZRDResult.ReadNextResult();
    BeamCounter = 1;                                             % Light beam counter
    RowCounter = 1;                                              % Row counter

    while success_NextResult == 1
        SegmentCounter = 1;                                      % Segment counter
        [t_success_NextSegmentFull, t_segmentLevel, t_segmentParent, t_hitObj, t_hitFace, t_insideOf, t_status, ...
              t_x, t_y, t_z, t_l, t_m, t_n, t_exr, t_exi, t_eyr, t_eyi, t_ezr, t_ezi, t_intensity, t_pathLength,...
              t_xybin, t_lmbin, t_xNorm, t_yNorm, t_zNorm, t_index, t_startingPhase, t_phaseOf, t_phaseAt] = ZRDResult.ReadNextSegmentFull();   

        while t_success_NextSegmentFull == 1
            [t_success_NextSegmentFull, t_segmentLevel, t_segmentParent, t_hitObj, t_hitFace, t_insideOf, t_status,...
            t_x, t_y, t_z, t_l, t_m, t_n, t_exr, t_exi, t_eyr, t_eyi, t_ezr, t_ezi, t_intensity, t_pathLength,...
            t_xybin, t_lmbin, t_xNorm, t_yNorm, t_zNorm, t_index, t_startingPhase, t_phaseOf, t_phaseAt] = ZRDResult.ReadNextSegmentFull(); 
            if t_success_NextSegmentFull == 1
                Beam(RowCounter,1) = BeamCounter;
                Segment(RowCounter,1) = SegmentCounter;
                SegmentLevel(RowCounter,1) = t_segmentLevel;
                SegmentParent(RowCounter,1) = t_segmentParent;
                hitObj(RowCounter,1) = t_hitObj;
                hitFace(RowCounter,1) = t_hitFace;
                insideOf(RowCounter,1) = t_insideOf;
                status{RowCounter,1} = string(t_status);
                x(RowCounter,1) = t_x;
                y(RowCounter,1) = t_y;
                z(RowCounter,1) = t_z;
                l(RowCounter,1) = t_l;     
                m(RowCounter,1) = t_m;     
                n(RowCounter,1) = t_n;  
                exr(RowCounter,1) = t_exr;
                exi(RowCounter,1) = t_exi;
                eyr(RowCounter,1) = t_eyr;
                eyi(RowCounter,1) = t_eyi;
                ezr(RowCounter,1) = t_ezr;
                ezi(RowCounter,1) = t_ezi;
                intensity(RowCounter,1) = t_intensity;
                pathLength(RowCounter,1) = t_pathLength;
                xybin(RowCounter,1) = t_xybin;
                lmbin(RowCounter,1) = t_lmbin;
                xNorm(RowCounter,1) = t_xNorm;
                yNorm(RowCounter,1) = t_yNorm;
                zNorm(RowCounter,1) = t_zNorm;
                index(RowCounter,1) = t_index;
                startingPhase(RowCounter,1) = t_startingPhase;
                phaseOf(RowCounter,1) = t_phaseOf;
                phaseAt(RowCounter,1) = t_phaseAt;

                SegmentCounter = SegmentCounter + 1;
                RowCounter = RowCounter + 1;    
            end            
        end

        % Read Next Beam
        [success_NextResult, rayNumber, waveIndex, wlUM, numSegments] = ZRDResult.ReadNextResult();
        BeamCounter = BeamCounter + 1;
    end

    RayTable = table(Beam, Segment, SegmentLevel, SegmentParent, hitObj, hitFace, insideOf, status,...
                 x, y, z, l, m, n, exr, exi, eyr, eyi, ezr, ezi, intensity, pathLength,...
                 xybin, lmbin, xNorm, yNorm, zNorm, index, startingPhase, phaseOf, phaseAt);

    if option
        if exist('RayData.xlsx')
            delete('RayData.xlsx');
        end
        writetable(RayTable, 'RayData.xlsx');
    end
end    