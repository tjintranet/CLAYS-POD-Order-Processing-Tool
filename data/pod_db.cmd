@echo off
setlocal

"C:\Program Files\EasyDataTransform_1_42_0\EasyDataTransform_1_42_0.exe" "\\192.168.10.169\pipeline3\clays_POD\data\Create Data JSON.transform" -file "PRH Master Titles=\\192.168.10.169\pipeline3\clays_POD\data\PRH Master Titles.xlsx[Sheet1]" -file "data1=\\192.168.10.169\pipeline3\clays_POD\data\data.json" -cli -verbose

rem
exit

