if not exist "C:\AutoExe\" mkdir C:\AutoExe
xcopy /H input\* C:\AutoExe
del /AH input\*

start AutoExcel.xlsm