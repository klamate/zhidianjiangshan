@echo off

set PublicTool=%cd%
pushd %PublicTool%
set PublicTool=%cd%
popd

set xlsx2csv=%PublicTool%\xlsx2csv.vbs

set ResourceDir=%cd%
pushd %ResourceDir%
set ResourceDir=%cd%
popd

set OutputDir=%cd%\..\..\res\res_config
pushd %OutputDir%
set OutputDir=%cd%
popd

echo %xlsx2csv%
echo %ResourceDir%
echo %OutputDir%

%xlsx2csv% "%ResourceDir%\建筑物.xlsx" "Sheet1" "%OutputDir%\building.csv"
%xlsx2csv% "%ResourceDir%\兵种.xlsx" "工作表1" "%OutputDir%\solider.csv"
%xlsx2csv% "%ResourceDir%\地图.xlsx" "地图1" "%OutputDir%\map1.csv"
pause