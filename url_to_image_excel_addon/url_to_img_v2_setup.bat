@ECHO OFF

REM |--------------------------------------------------------------------------------------------------------------------
REM | Purpose:  Generic Excel Addin Install
REM |--------------------------------------------------------------------------------------------------------------------


REM
REM     /E   = Copies directories and sub-directories, including empty ones. Same as /S /E. May be used to modify /T. 
REM     /D:m-d-y = Copies files changed on or after the specified date. 
REM        If no date is given, copies only those files whose source time is newer than the destination time. 
REM     /K   = Copies attributes. Normal Xcopy will reset read-only attributes. 
REM     /Q   = Does not display file names while copying. 
REM     /R   = Overwrites read-only files. 
REM     /Y   = Suppresses prompting to confirm you want to overwrite an existing destination file. 
REM

REM Copy the install directory and sub-directories
REM echo f | XCOPY ".\url_to_img_macro.xlam" "%AppData%\Microsoft\AddIns\url_to_img_macro.xlam" /E /K /Q /R /Y /D
    echo f | XCOPY ".\url_to_img_macro_v2.xlam" "%AppData%\Microsoft\Excel\XLSTART\url_to_img_macro.xlam" /E /K /Q /R /Y /D
REM echo f | XCOPY ".\url_to_img_macro.xlam" "%AppData%\Roaming\Microsoft\Excel\XLSTART\url_to_img_macro.xlam" /E /K /Q /R /Y /D


