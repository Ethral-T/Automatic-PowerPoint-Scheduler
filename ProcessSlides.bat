@echo off
setlocal enabledelayedexpansion
title Auto Schedule Processor

:: Deletes old files from remote access locations
echo Deleting old files from remote access locations...
call :DeleteTask "%cd%\Room Assignments"

:: Run the VBScript to convert PowerPoint slides to 4K JPEGs
echo Processing PowerPoint files...
cscript //nologo ConvertSlides.vbs "%cd%\Room Assignments"

:: Rename new images in the Room Assignments folder
echo Renaming Images...
pushd "%cd%\Room Assignments" && (
    set "day[1]=Monday"
    set "day[2]=Tuesday"
    set "day[3]=Wednesday"
    set "day[4]=Thursday"
    set "day[5]=Friday"
    set "day[6]=Saturday"
    set "day[7]=Sunday"
    for /L %%G in (1,1,7) do (
        ren "000%%G_Slide%%G.jpg" "RA%%G_!day[%%G]!.jpg"
    )
popd
) ||(
    echo Room Assignments directory not found.
)

:: Copy the new images to the remote access locations
echo Copying new images to remote access locations...
xcopy "%cd%\Room Assignments" "C:\Users\%username%\Desktop\Room Assignments\" /I /Y

:: Copy USB Contents to backup location
echo Copying USB contents to backup location...
xcopy "%cd%\*" "C:\Users\%username%\Desktop\TV USB Stick Backup" /I /Y

:: Close PowerPoint application in event of a failure in the VBScript
echo Closing PowerPoint application...
taskkill /f /im powerpnt.exe

endlocal

:: Define a subroutine to delete files from a directory that can be called upon later
:DeleteTask
set "folder=%~1"
pushd "%folder%" && (
    del /Q *
    popd
) ||(
    echo %folder% directory not found.
)
goto :eof