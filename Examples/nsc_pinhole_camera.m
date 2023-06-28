clear all;
close all;
clc;

cd ..

ZOS = ZDDE();
TheApplication = ZOS.TheApplication;
TheApplication.PrimarySystem.New(false);                  % 新建文件
TheNCE = TheApplication.PrimarySystem.NCE;
TheSystem = TheApplication.PrimarySystem;
TheSystem.MakeNonSequential();                            % 设置为非序列模式

%% 矩形光源
TheNCE.InsertNewObjectAt(1);                              % 插入新物体
SourceRectangle = TheNCE.GetObjectAt(1);     
ObjectType_SR = SourceRectangle.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.SourceRectangle);
SourceRectangle.ChangeType(ObjectType_SR);                % 获取非序列组件的第1个物体作为矩形光源
SourceRectangle.Comment = '矩形光源';                     % 设置备注
SourceRectangle.ObjectData.XHalfWidth = 2.5;              % 设置 x 半宽
SourceRectangle.ObjectData.YHalfWidth = 2.5;              % 设置 y 半宽
SourceRectangle.ObjectData.NumberOfLayoutRays = 2e4;      % 设置阵列光线条数
SourceRectangle.ObjectData.NumberOfAnalysisRays = 1e9;    % 设置分析光线条数
SourceRectangle.SourcesData.SourceColor = ZOSAPI.Editors.NCE.SourceColorMode.D65White;  % 设置光源颜色为 D65 白光

%% 幻灯片位置面
TheNCE.InsertNewObjectAt(2);                              % 插入新物体
SlideSurface = TheNCE.GetObjectAt(2);
ObjectType_R = SlideSurface.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.Rectangle);
SlideSurface.ChangeType(ObjectType_R);                    % 获取非序列组件的第2个物体作为矩形
SlideSurface.Comment = '幻灯片基片';                      % 设置备注
SlideSurface.ObjectData.XHalfWidth = 2.5;                 % 设置 x 半宽
SlideSurface.ObjectData.YHalfWidth = 2.5;                 % 设置 y 半宽
SlideSurface.ZPosition = 1.0;                             % Z 位置

%% 插入幻灯片
TheNCE.InsertNewObjectAt(3);                              % 插入新物体
Slide = TheNCE.GetObjectAt(3);
ObjectType_SL = Slide.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.Slide);
Slide.ChangeType(ObjectType_SL);                          % 获取非序列组件的第3个物体作为幻灯片
Slide.Comment = '幻灯片';                                 % 设置幻灯片
Slide.RefObject = 2;                                      % 设置参考物体
Slide.ZPosition = 1e-2;                                   % 设置 Z 位置
Slide.ObjectData.XFullWidth = 2.5;                        % 设置 x 全宽
Slide.Comment = 'Alex200.BMP';                            % 设置注释（幻灯片内容）

%% 插入小孔
TheNCE.InsertNewObjectAt(4);                              % 插入新物体
Pinhole = TheNCE.GetObjectAt(4);
ObjectType_AN = Slide.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.Annulus);
Pinhole.ChangeType(ObjectType_AN);                        % 获取非序列组件的第 4 个物体作为环形面
Pinhole.Comment = '小孔';                                 % 设置幻灯片
Pinhole.RefObject = 3;                                    % 设置参考物体
Pinhole.ZPosition = 10;                                   % 设置 Z 位置
Pinhole.Material = 'ABSORB';                              % 设置材料
Pinhole.ObjectData.MaxXHalfWidth = 9.0;                   % 设置环形面外径
Pinhole.ObjectData.MaxYHalfWidth = 9.0;
Pinhole.ObjectData.MinXHalfWidth = 0.05;                  % 设置环形面内径
Pinhole.ObjectData.MinYHalfWidth = 0.05;

%% 插入彩色探测器
TheNCE.InsertNewObjectAt(5);                              % 插入新物体
Detector = TheNCE.GetObjectAt(5);
ObjectType_DT = Slide.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.DetectorColor);
Detector.ChangeType(ObjectType_DT);
Detector.Comment = '彩色探测器';                          % 设置彩色探测器
Detector.Material = 'ABSORB';                             % 设置材料
Detector.ZPosition = 25;                                  % 设置 Z 位置
Detector.RefObject = 4;                                   % 设置参考物体
Detector.ObjectData.NumberXPixels = 500;
Detector.ObjectData.NumberYPixels = 500;
Detector.ObjectData.XHalfWidth = 6.25;
Detector.ObjectData.YHalfWidth = 6.25;
Detector.ObjectData.Color = 4;                            % 设置探测器颜色设置为4，即真彩色

%% 设置幻灯片位置面散射
% 设置散射类型
o3_Scatter = SlideSurface.CoatScatterData.GetFaceData(0).CreateScatterModelSettings(ZOSAPI.Editors.NCE.ObjectScatteringTypes.Lambertian);
o3_Scatter.S_Lambertian_.ScatterFraction = 1.0;
SlideSurface.CoatScatterData.GetFaceData(0).ChangeScatterModelSettings(o3_Scatter);
SlideSurface.CoatScatterData.GetFaceData(0).NumberOfRays = 1;
% 设置散射路径
SlideSurface.ScatterToData.ScatterToMethod = ZOSAPI.Editors.NCE.ScatterToType.ImportanceSampling; % 散射路径模型 --> 重点采样
% 设置重点采样数据
ImportanceData = SlideSurface.ScatterToData.GetRayData(1);
ImportanceData.Towards = 4;
ImportanceData.Size = 0.4;
ImportanceData.Limit = 1;
SlideSurface.ScatterToData.SetRayData(1,ImportanceData);

%% 光束追迹
NSCRayTrace = TheSystem.Tools.OpenNSCRayTrace();
NSCRayTrace.SplitNSCRays = false;
NSCRayTrace.ScatterNSCRays = true;
NSCRayTrace.UsePolarization = false;
NSCRayTrace.IgnoreErrors = true;
NSCRayTrace.SaveRays = false;
NSCRayTrace.ClearDetectors(0);
NSCRayTrace.RunAndWaitForCompletion();
NSCRayTrace.Close();

%% 打开探测查看器
analysis = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.DetectorViewer);


%% 保存文件
TheApplication.PrimarySystem.SaveAs(fullfile(pwd, '\zmx files\Short course\Fundamentals of Optics\nsc_pinhole_camera.zmx'));