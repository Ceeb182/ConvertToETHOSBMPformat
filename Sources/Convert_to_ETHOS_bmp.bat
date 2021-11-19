@echo off
REM ######################################################################### 
REM #                                                                       #
REM # Copyright (C) Ceeb182@laposte.net                                     #
REM #                                                                       #
REM # License GPLv2: http://www.gnu.org/licenses/gpl-2.0.html               #
REM #                                                                       #
REM # This program is free software; you can redistribute it and/or modify  #
REM # it under the terms of the GNU General Public License version 2 as     #
REM # published by the Free Software Foundation.                            #
REM #                                                                       #
REM # This program is distributed in the hope that it will be useful        #
REM # but WITHOUT ANY WARRANTY; without even the implied warranty of        #
REM # MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the         #
REM # GNU General Public License for more details.                          #
REM #                                                                       #
REM ######################################################################### 
setlocal DisableDelayedExpansion
set ResultDirName=FrSkyEthosBMP
set Mypath=%cd%\%ResultDirName%\
set cnt=0
set ValidFile=false
set ImageMagickExist=false
set OnlyConvert=true
set MyChoice=e

if "%~1" == "" (
	set OnlyConvert=true
	set MyChoice=c
) else (
	if "%~1" == "-m" (
		set OnlyConvert=false
	) else (
		echo Unknown command for this script
		echo Use -m as argument to display script's menu
		exit /b
	)
)

echo *********************************************************
echo * Script for Windows OS to convert any type of image    *
echo * (regardless of file format, size and internal coding) *
echo * to the BMP format used by the FrSky Ethos OS          *
echo *                                                       *
echo * 32bits BMP format / 8 bits per colour + Alpha channel *
echo *                   size : 300x280px                    *
echo *                                                       *
echo * Version 1.0                                by Ceeb182 *
echo *********************************************************
echo *
rem >> Test if ImageMagick exist
magick -version >nul 2>&1 && (set ImageMagickExist=true) 
if 	%ImageMagickExist%==true (
	echo * Program ImageMagick found !
) else (
	echo * Please install ImageMagick at https://www.imagemagick.org
	echo * to use this script.
	echo *
	set /p DUMMY=* Hit ENTER to continue...
	exit /b
)
rem >> Command list
if %OnlyConvert%==false (
	echo *
	echo * Choose your command :
	echo *   c  : to Convert pictures to BMP format for FrSky Ethos OS
	echo *   d  : to Delete a previous result
	set /p MyChoice=* Enter your choice :
)

if %MyChoice%==c (
	echo *
	echo * COMMAND : Convert
    rem >> Check and count files to convert
	setlocal EnableDelayedExpansion
	for %%f in (*.*) do (
		set ValidFile=false
		for %%g in (.jpg,.jpeg,.png,.bmp,.gif,.webp,.svg,.tiff,.ico) do if %%g==%%~xf set ValidFile=yes
		if !ValidFile!==yes set /a cnt+=1
	)	
	rem >> Create directory for results
	if not !cnt!==0 if not exist "%Mypath%" (	
		mkdir "%Mypath%"
		echo * Create %ResultDirName% directory : DONE !
	)
	rem >> Convert all picture	
	echo * Number of image files to convert = !cnt!
	set cnt=0
    for %%f in (*.*) do (
		set ValidFile=false
		for %%g in (.jpg,.jpeg,.png,.bmp,.gif,.webp,.svg,.tiff,.ico) do if %%g==%%~xf set ValidFile=yes
		if !ValidFile!==yes (
			set /a cnt+=1
			magick "%%f" -resize 300x280 -alpha Set -depth 8 -compose Copy -gravity center -background none -extent 300x280 "%Mypath%%%~nf.bmp"
			echo * Conversion #!cnt! : %%f
		)
	)
	echo *
	if not !cnt!==0 echo * Check results in %Mypath%
	endlocal
	set /p DUMMY=* Hit ENTER to continue...
	exit /b
)
if %MyChoice%==d (
	echo *
	echo * COMMAND : Delete a previous result 
	if exist "%Mypath%" (
		echo * %ResultDirName% directory found.
		rmdir /q /s "%Mypath%"
		echo * Delete a previous result : DONE !
	) else (
		echo * %ResultDirName% directory not found. Nothing to delete.
	)
	echo *
	set /p DUMMY=* Hit ENTER to continue...
	exit /b
)