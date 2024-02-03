@echo off
setlocal enabledelayedexpansion
title Auto Scheduler written by Taylor

echo Deleting old files from remote access locations...
:: Deletes old files from remote access locations
call :DeleteTask "%cd%~\Room Assignments\"

:: Run the VBScript to convert PowerPoint slides to 4K JPEGs
echo Processing PowerPoint files...
cscript //nologo ConvertSlides.vbs "%cd%\Room Assignments"

:: Rename new images in the Room Assignments folder
echo Renaming Images...
pushd "%cd%~\Room Assignments" && (
    for /L %%G in (1,1,7) do (
        set "day=Extra"
        if %%G==1 set "day=Monday"
        if %%G==2 set "day=Tuesday"
        if %%G==3 set "day=Wednesday"
        if %%G==4 set "day=Thursday"
        if %%G==5 set "day=Friday"
        if %%G==6 set "day=Saturday"
        if %%G==7 set "day=Sunday"
        ren "000%%G_Slide%%G.jpg" "RA%%G_!day!.jpg"
    )
popd
) ||(
    echo Room Assignments directory not found.
)

:: Copy the new images to the remote access locations
echo Copying new images to remote access locations...
xcopy "%cd%~\Room Assignments" "C:\Users\%username%\Desktop\Room Assignments\" /I /Y

:: Copy USB Contents to backup location
echo Copying USB contents to backup location...
xcopy "%cd%\*" "C:\Users\%username%\Desktop\TV USB Stick Backup" /I /Y

:: Close PowerPoint application in event of a failure in the VBScript
echo Closing PowerPoint application...
taskkill /f /im powerpnt.exe

endlocal

:DeleteTask
set "folder=%~1"
pushd "%folder%" && (
    del /Q *
    popd
) ||(
    echo %folder% directory not found.
)
goto :eof