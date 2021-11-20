; ######################################################################### 
; #                                                                       #
; # Copyright (C) Ceeb182@laposte.net                                     #
; #                                                                       #
; # License GPLv2: http://www.gnu.org/licenses/gpl-2.0.html               #
; #                                                                       #
; # This program is free software; you can redistribute it and/or modify  #
; # it under the terms of the GNU General Public License version 2 as     #
; # published by the Free Software Foundation.                            #
; #                                                                       #
; # This program is distributed in the hope that it will be useful        #
; # but WITHOUT ANY WARRANTY; without even the implied warranty of        #
; # MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the         #
; # GNU General Public License for more details.                          #
; #                                                                       #
; ######################################################################### 
; # Coded with AutoIt v3   /   https://www.autoitscript.com/              #
; #########################################################################

#include <GUIConstantsEx.au3>
#include <ButtonConstants.au3>
#include <FontConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <UpdownConstants.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <AutoItConstants.au3>
#include <StringConstants.au3>


Dim Const $sVersion="v1.1"
; MAIN GUI (Design)
Dim $iGui, $iGUIWidth = 500, $iGUIHeight = 200, $aParentWin_Pos
Dim $SpaceLb1ChB1=20, $SpaceBtwChB=25
Dim $idLabel1, $Label1Margin = 10, $Label1Height=50; Label
Dim $idChkBox1, $idChkBox2, $idChkBox3, $idProgressBar
Dim $idConvert_Btn, $idCancel_Btn, $idPreferences_Btn, $iMsg
; PREFERENCE GUI (Design)
Dim $idGuiPref, $iGuiPrefWidth = 500, $iGuiPrefHeight = 310, $iIcon=32, $iGuiPregYPos
Dim $idGuiPrefLabelInfo, $idGuiPrefIcon
Dim $idGuiPrefGroupRename, $idGuiPrefGroupTransfer
Dim $idGuiPrefUppercase, $idGuiPrefCountChar, $idGuiPrefUppercase, $idGuiPrefSpecialChar, $idGuiCharSet
Dim $idGuiPrefOverwrite, $idGuiPrefDontTransfer, $idGuiPrefOk, $idGuiPrefCharNumber, $idGuiPrefDefault

; Preferences parameters
; - Default
Dim Const $bDAutorename=True, $bDOverwriteComputer=True, $bDTransferSD=True
Dim Const $bDAvoidUppercase=True, $bDLimitLength=True, $iDLimitLength=11, $iDLimitLengthMax=64, $iDLimitLengthMin=6
Dim Const $bDLimitSpecialChar=True, $sDSpecialChar=";[]+=-_", $sDSpecialCharSet=" ;[]+=()!-_@#", $sMandatoryChar=" ;[]+="
Dim Const $bDOverwriteSD=True
; - Current (set to default on startup before loading INI file)
Dim $bAutorename=$bDAutorename, $bOverwriteComputer=$bDOverwriteComputer, $bTransferSD=$bDTransferSD
Dim $bAvoidUppercase=$bDAvoidUppercase, $bLimitLength=$bDLimitLength, $iLimitLength=$iDLimitLength
Dim $bLimitSpecialChar=$bDLimitSpecialChar, $sSpecialChar=$sDSpecialChar
Dim $bOverwriteSD=$bDOverwriteSD
; - Ini File
Dim $sIniFile				 ; Filename of INI file to save Preferences

; Other variables
Dim $iPID, $sOutput 		 ; To execute DOS command
Dim $Value					 ; Dummy value
Dim $RunMode, $PathDvpt 	 ; Debug parameters
Dim $sFile, $s, $sFileList	 ; To find image files
Dim $MyPath					 ; Path of current Script
Dim $sFolder 				 ; Full path for output on computer
Dim $iFileImageNumber		 ; Number of image files matching the allowed extensions
Dim $sEthosSDDrive			 ; Ethos SD Card drive
Dim $sSuffix				 ; Possible suffix given to the output file name to avoid duplicates
Dim Const $sBaseSuffix="abcdefghijklmnopqrstuvwxyz" ; Character base used to generate a suffix
Dim $aFilenameAfterConvert[0]; List of filename after converting to BMP (used to transfer to SD Card)
Dim $bConvertInProgress      ; Use to disable DoubleClick when converting
Dim $ImageExt[] = [".jpg",".jpeg",".png",".bmp",".gif",".webp",".svg",".tiff",".ico"] ; allowed extension
Dim Const $sOutputFolder="FrSkyEthosBMP" ; Name of output directory on computer


_Main()

Func _Main()
	Local $iValue, $sString
	$oMyError = ObjEvent("AutoIt.Error","MyErrFunc")
	$MyPath =@WorkingDir & "\"
	$sIniFile=GetFileName(@ScriptName) & ".ini" ; INI file has the same name as current script name
	$bConvertInProgress=False
	
	; >> Test if "ImageMagick" program exist on the computer
	; >> ImageMagick is the program that converts images
	$iPID = Run(@ComSpec & " /C magick -version >nul 2>&1 && (echo true) || (echo false)","",@SW_HIDE, $STDOUT_CHILD)
	ProcessWaitClose($iPID)
	$sOutput = StringTrimRight(StringStripCR(StdoutRead($iPID)), StringLen(@CRLF) -1 )
	If $sOutput="false" Then ; ImageMagick doesn't exist on your computer
		$Value = MsgBox ($MB_ICONERROR+$MB_OKCANCEL+$MB_DEFBUTTON2, "Unable to find ImageMagick", _
						"You must install ImageMagick on your computer" & @CRLF & _
						"to run this converter." & @CRLF & _
						"Get this program from https://imagemagick.com/" & @CRLF & _
						"Do you want to follow this weblink ?")
		If $Value=$IDOK Then ShellExecute("https://imagemagick.com/")
		Exit
	EndIf
	; >> Count picture files to convert
	CountPictureFileToConvert()
	If $iFileImageNumber=0 Then
		MsgBox($MB_OK+$MB_ICONINFORMATION,"Nothing to do here!","In the current directory, I did not find any image files to convert.")
		Exit
	EndIf
	; >> Load preferences (Default or INI file setting if it exists)
	LoadPreferences()
	; >> Gui creation
	GuiMain()
	GuiPreference()
	; >> Looking for Ethos SD card
	FindEthosSD()
	; >> Set all Gui to match current preferences (default if INI file doesn't exist)
	SetPreferenceOnGUI()
	
	GUIRegisterMsg($WM_COMMAND, "DoubleClick")
	; >> Show window/Make the window visible
	GUISetState(@SW_SHOW,$iGui)

	While 1 ; Loop until: Esc key OR Alt+F4 OR user clicks on CANCEL or CONVERT button
		$iMsg = GUIGetMsg(1)
		Switch $iMsg[1] 
			Case $iGui ; If it's Main GUI
				Switch $iMsg[0]
					Case $GUI_EVENT_CLOSE, $idCancel_Btn ; Close button OR CANCEL button
						SavePreferences()
						GUIDelete($idGuiPref)
						GUIDelete($iGui)
						Exit
					Case $idChkBox1 ; CheckBox 'Automatically rename the output files'
						(GUICtrlRead($idChkBox1)=$GUI_CHECKED) ? (SetBool($bAutorename)) : (UnsetBool($bAutorename))
					Case $idChkBox2 ; CheckBox 'Overwrite if output file exist on computer'
						(GUICtrlRead($idChkBox2)=$GUI_CHECKED) ? (SetBool($bOverwriteComputer)) : (UnsetBool($bOverwriteComputer))
					Case $idChkBox3 ; CheckBox 'Transfer BMP files to the SD Card'
						(GUICtrlRead($idChkBox3)=$GUI_CHECKED) ? (SetBool($bTransferSD)) : (UnsetBool($bTransferSD))
						If $bTransferSD Then ; Change text on Convert button
							GUICtrlSetData($idConvert_Btn,"Convert and transfer")
						Else
							GUICtrlSetData($idConvert_Btn,"Convert")
						Endif
					Case $idConvert_Btn ; CONVERT button
						$bConvertInProgress=True
						ConvertToEthosBMPFormat()
						If $bTransferSD=True And $sEthosSDDrive<>"" Then
							TransferToSDCard()
						EndIf
						SavePreferences()
						; Message to warn the user where the results are
						GUICtrlSetTip($idProgressBar, "Results are in " & $MyPath & $sOutputFolder)
						For $iValue=1 to 4 step 1
							GUICtrlSetData($idLabel1,"")
							Sleep(200)
							GUICtrlSetData($idLabel1,">>> Results are in local """ & $sOutputFolder & """ directory <<<")
							Sleep(600)
						Next
						; Delete all GUIs
						GUIDelete($idGuiPref)
						GUIDelete($iGui)
						Exit
					Case $idPreferences_Btn ; PREFERENCES button
						GUISetState(@SW_DISABLE,$iGui)
						GUICtrlSetState($idPreferences_Btn, $GUI_DISABLE)
						GUICtrlSetState($idCancel_Btn, $GUI_DISABLE)
						GUICtrlSetState($idConvert_Btn, $GUI_DISABLE)
						GUISetState(@SW_SHOW,$idGuiPref)
				EndSwitch
			Case $idGuiPref ; If it's PREFERENCES Gui
				Switch $iMsg[0]
					Case $GUI_EVENT_CLOSE, $idGuiPrefOk ; Close button OR Ok button
						GUISetState(@SW_HIDE,$idGuiPref)
						GUISetState(@SW_ENABLE,$iGui)
						WinActivate($iGui)
						GUICtrlSetState($idPreferences_Btn, $GUI_ENABLE)
						GUICtrlSetState($idCancel_Btn, $GUI_ENABLE)
						GUICtrlSetState($idConvert_Btn, $GUI_ENABLE)
					Case $idGuiPrefIcon ; Logo of the script
						ShellExecute("https://github.com/Ceeb182/ConvertToETHOSBMPformat")
					Case $idGuiPrefUppercase ; CheckBox 'Avoid uppercase only'
						(GUICtrlRead($idGuiPrefUppercase)=$GUI_CHECKED) ? (SetBool($bAvoidUppercase)) : (UnsetBool($bAvoidUppercase))
					Case $idGuiPrefCountChar ; CheckBox 'Limit the name to X characters'
						(GUICtrlRead($idGuiPrefCountChar)=$GUI_CHECKED) ? (SetBool($bLimitLength)) : (UnsetBool($bLimitLength))
					Case $idGuiPrefCharNumber ; UpDown control to set max length name
						$iLimitLength = GUICtrlRead($idGuiPrefCharNumber)
						If $iLimitLength<$iDLimitLengthMin Then $iLimitLength=$iDLimitLengthMin
						If $iLimitLength>$iDLimitLengthMax Then $iLimitLength=$iDLimitLengthMax
						GUICtrlSetData($idGuiPrefCharNumber,$iLimitLength)
						GUICtrlSetData($idGuiPrefCountChar,"Limit the name to " & $iLimitLength & " characters")
					Case $idGuiPrefSpecialChar ; CheckBox 'Allow only this special character set'
						(GUICtrlRead($idGuiPrefSpecialChar)=$GUI_CHECKED) ? (SetBool($bLimitSpecialChar)) : (UnsetBool($bLimitSpecialChar))
					Case $idGuiCharSet ; Input to set string of special character
						$sSpecialChar = CleanSpecialChar(GUICtrlRead($idGuiCharSet))
						GUICtrlSetData($idGuiCharSet,$sSpecialChar)
					Case $idGuiPrefOverwrite ; Radio button 'Overwrite on SD'
						(GUICtrlRead($idGuiPrefOverwrite)=$GUI_CHECKED) ? (SetBool($bOverwriteSD)) : (UnsetBool($bOverwriteSD))
					Case $idGuiPrefDontTransfer ; Radio button 'Do not transfer this file'
						(GUICtrlRead($idGuiPrefDontTransfer)=$GUI_UNCHECKED) ? (SetBool($bOverwriteSD)) : (UnsetBool($bOverwriteSD))
					Case $idGuiPrefDefault ; Default button
						$bAutorename=$bDAutorename
						$bOverwriteComputer=$bDOverwriteComputer
						$bTransferSD=$bDTransferSD
						$bAvoidUppercase=$bDAvoidUppercase
						$bLimitLength=$bDLimitLength
						$iLimitLength=$iDLimitLength
						$bLimitSpecialChar=$bDLimitSpecialChar
						$sSpecialChar=$sDSpecialChar
						$bOverwriteSD=$bDOverwriteSD
						SetPreferenceOnGUI()
				EndSwitch
		EndSwitch
	WEnd
EndFunc   
; --------------------|End of Main|------------------------

Func TransferToSDCard()
	; >> Transfert to Ethos SD Card
	Local $iOuputFileNumber, $i
	FindEthosSD() ; Looking for Ethos SD card
	If $sEthosSDDrive="" Then Return
	$iOuputFileNumber=UBound($aFilenameAfterConvert)
	; Update GUI
	GUICtrlSetTip($idProgressBar, "Tranfer to ETHOS SD Card") ; Update ProgressBarTip
	GUICtrlSetData($idLabel1,"")     ; Update label under ProgressBar
	GUICtrlSetData($idProgressBar, 0); Update ProgressBar
	For $i=0 To $iOuputFileNumber-1 Step 1
		GUICtrlSetTip($idProgressBar, "Tranfer to ETHOS SD Card " & $i & "/" & $iFileImageNumber) ; Update ProgressBarTip
		If FileExists($MyPath & $sOutputFolder & "\" & $aFilenameAfterConvert[$i] &".bmp") Then
			If $bOverwriteSD=False And FileExists($sEthosSDDrive & "\bitmaps\user\" & $aFilenameAfterConvert[$i] &".bmp") Then
				GUICtrlSetData($idLabel1,"Transfer to Ehos SD [" & $i & "/" & $iFileImageNumber & "] > Skip file : " & $aFilenameAfterConvert[$i] &".bmp")
			Else
				GUICtrlSetData($idLabel1,"Transfer to Ehos SD [" & $i & "/" & $iFileImageNumber & "] > " & $aFilenameAfterConvert[$i] &".bmp")
				FileCopy($MyPath & $sOutputFolder & "\" & $aFilenameAfterConvert[$i] &".bmp", _
					$sEthosSDDrive & "\bitmaps\user\" & $aFilenameAfterConvert[$i] &".bmp", $FC_OVERWRITE )
			EndIf
		Else
			GUICtrlSetData($idLabel1,"Transfer to Ehos SD [" & $i & "/" & $iFileImageNumber & "] > Unable to find the file to transfer !")
		EndIf
		GUICtrlSetData($idProgressBar, Int(($i+1)/$iOuputFileNumber*100)); Update ProgressBar
	Next
EndFunc

Func ConvertToEthosBMPFormat()
	; >> Convert image files to Ethos BMP format on computer
	Local $sRawFileNameOutput, $sFinalFileNameOutput
	Local $cCurrentCharInName, $cCurrentSpecialChar, $sAtLeastCharset
	Local $bRules,$iSuffixNumber
	Local $iOuputFileNumber
	; >> Modify main Gui with progress bar
	GuiTransfer()
	; >> Create output directory
	$sFolder = $MyPath & $sOutputFolder
	If Not FileExists($sFolder) Then
		$Value = DirCreate($sFolder)
		If $Value=0 Then
			MsgBox ($MB_ICONERROR+$MB_OK, "Something is wrong !", "Unable to create output folder in current directory")
			Exit
		EndIf
	Else
		;MsgBox ($MB_SYSTEMMODAL,"Debug","Folder" & $sOutputFolder & " exist !") ; Debug
	EndIf


	; >> Convert image files to FrSky Ethos BMP format
	$iCurrentFileImageNumber=0
	$s = FileFindFirstFile($MyPath & '*')
	If $s <> -1 Then
		While 1
			$sFile = FileFindNextFile($s)
			If @error Then ExitLoop
			If Not @extended Then
				If IsPictureFile($sFile) Then 
					$iCurrentFileImageNumber +=1
					$sRawFileNameOutput=GetFileName($sFile)
					$sFinalFileNameOutput=""
					; >> Autorename (3 possible operations : Special charset, control lenght, avoir uppercase
					If $bAutorename=True Then ; Autorename enabled
						; ## Comply with ETHOS naming rules ##
						; #> Step 1/3 : Forbid accented characters like é, è, ç, ù
						; #> Step 2/3 : Check if charset is : A~Z, a~z, 0~9, ()!-_@#;[]+= `Space`
						; #> Step 3/3 : At least one of these characters in the name : a~z or ;[]+=`Space` (Not here)
						If $bAvoidUppercase=True Then
							For $i=1 to StringLen($sRawFileNameOutput) Step 1
								$cCurrentCharInName=StringMid($sRawFileNameOutput,$i,1)
								; Check if it's A~Z a~z 0~9
								If (Asc($cCurrentCharInName)>64 AND Asc($cCurrentCharInName)<91) OR _
									(Asc($cCurrentCharInName)>96 AND Asc($cCurrentCharInName)<123) OR _
									(Asc($cCurrentCharInName)>47 AND Asc($cCurrentCharInName)<58) Then
									$sFinalFileNameOutput=$sFinalFileNameOutput & $cCurrentCharInName
								EndIf
								If $cCurrentCharInName="é" OR $cCurrentCharInName="è" OR $cCurrentCharInName="ê" OR $cCurrentCharInName="ë" Then _
											$sFinalFileNameOutput=$sFinalFileNameOutput & "e"
								If $cCurrentCharInName="ô" OR $cCurrentCharInName="ö" Then _
											$sFinalFileNameOutput=$sFinalFileNameOutput & "o"
								If $cCurrentCharInName="î" OR $cCurrentCharInName="ï" Then _
											$sFinalFileNameOutput=$sFinalFileNameOutput & "i"
								If $cCurrentCharInName="â" OR $cCurrentCharInName="ä" Then _
											$sFinalFileNameOutput=$sFinalFileNameOutput & "a"
								If $cCurrentCharInName="ù" OR $cCurrentCharInName="û" OR $cCurrentCharInName="ü" Then _
											$sFinalFileNameOutput=$sFinalFileNameOutput & "u"
							
								; Limit special charset to ()!-_@#;[]+=`Space`						
								If StringInStr($sDSpecialCharSet,$cCurrentCharInName)>0 Then $sFinalFileNameOutput=$sFinalFileNameOutput & $cCurrentCharInName
							Next
							If StringLen($sFinalFileNameOutput)=0 Then 
								$sFinalFileNameOutput="Image" ; Set a default name if name become empty
							EndIf
						Else
							$sFinalFileNameOutput=$sRawFileNameOutput
						EndIf
						; ## Custom limit special charset ##
						If $bLimitSpecialChar=True Then	; Checks if it is an allowed special character
							For $i=1 to StringLen($sRawFileNameOutput) Step 1
								$cCurrentCharInName=StringMid($sFinalFileNameOutput,$i,1)
								If StringInStr($sDSpecialCharSet,$cCurrentCharInName)>0 Then ; It's a special char
									If StringInStr($sSpecialChar,$cCurrentCharInName)=0 Then ; But it is not part of the special char custom list
										$sFinalFileNameOutput=StringLeft($sFinalFileNameOutput,$i-1) & StringMid($sFinalFileNameOutput,$i+1)
										$i=$i-1
									EndIf
								EndIf
							Next
							If StringLen($sFinalFileNameOutput)=0 Then 
								$sFinalFileNameOutput="Image" ; Set a default name if name become empty
							EndIf
						EndIf
						; ## Control length ##
						If $bLimitLength=True AND StringLen($sFinalFileNameOutput)>$iLimitLength Then ; Check if lenght control is enabled
							$sFinalFileNameOutput=StringMid($sFinalFileNameOutput,1,$iLimitLength)	
						Endif
						$sFinalFileNameOutput=StringStripWS ($sFinalFileNameOutput, $STR_STRIPTRAILING) ; Delete whitespace at the end of the name
						; ## Comply with ETHOS naming rules ##
						; #> Step 3/3 : At least one of these characters in the name : a~z or $sMandatoryChar=" ;[]+="
						If $bAvoidUppercase=True And CheckEthosNamingRules($sFinalFileNameOutput)=0 Then ; Lack one mandatory char
							If StringLen($sFinalFileNameOutput)+1<=$iLimitLength Then
								$sFinalFileNameOutput=$sFinalFileNameOutput & "a"
							Else
								$sFinalFileNameOutput=StringLeft($sFinalFileNameOutput,StringLen($sFinalFileNameOutput)-1) & "a"
							EndIf
						EndIf
						
						; ## Save filename in base ##
						ReDim $aFilenameAfterConvert[UBound($aFilenameAfterConvert)+1]
						$iOuputFileNumber=UBound($aFilenameAfterConvert)
						$aFilenameAfterConvert[$iOuputFileNumber-1]=$sFinalFileNameOutput
						
						; ## Check if two files don't have the same name ##
						$iSuffixNumber=0
						While FilenameExistsInBase()
							$iSuffixNumber+=1
							SetSuffix($iSuffixNumber,$sBaseSuffix) ; Generate suffix in $sSuffix
							If StringLen($sFinalFileNameOutput)+StringLen($sSuffix)<=$iLimitLength Then
								$aFilenameAfterConvert[$iOuputFileNumber-1]=$sFinalFileNameOutput & $sSuffix
							Else
								If StringLen($sSuffix)<StringLen($sFinalFileNameOutput) Then
									$aFilenameAfterConvert[$iOuputFileNumber-1]=StringLeft($sFinalFileNameOutput,StringLen($sFinalFileNameOutput)-StringLen($sSuffix)) & $sSuffix
								Else
									$sFinalFileNameOutput="Image" ; Set a default name if name become empty
									$aFilenameAfterConvert[$iOuputFileNumber-1]="Image"
									$iSuffixNumber=0
								EndIf
							Endif
						Wend
						$sFinalFileNameOutput=$aFilenameAfterConvert[$iOuputFileNumber-1]
					Else
						$sFinalFileNameOutput=$sRawFileNameOutput
					EndIf
					
					; >> Overwrite on computer
					GUICtrlSetTip($idProgressBar, "Convert " & $iCurrentFileImageNumber & "/" & $iFileImageNumber) ; Update ProgressBarTip
					If $bOverwriteComputer=False And FileExists($MyPath & $sOutputFolder & "\" & $sFinalFileNameOutput &".bmp") Then
						GUICtrlSetData($idLabel1,"Convert " & $iCurrentFileImageNumber & "/" & $iFileImageNumber & " > Skip """& $sFile & """  (" & $sFinalFileNameOutput & ".bmp)")
					Else
						GUICtrlSetData($idLabel1,"Convert " & $iCurrentFileImageNumber & "/" & $iFileImageNumber & " > """& $sFile & """  to  """ & $sFinalFileNameOutput & ".bmp""")
						$iPID = Run(@ComSpec & " /C magick """ & $sFile & "[0]"" -resize 300x280 -alpha Set -depth 8 -compose Copy -gravity center -background none -extent 300x280 """ & $MyPath & $sOutputFolder & "\" & $sFinalFileNameOutput &".bmp""","",@SW_HIDE, $STDERR_MERGED)
						ProcessWaitClose($iPID)
						$sOutput = StringTrimRight(StringStripCR(StdoutRead($iPID)), StringLen(@CRLF) -1 )
						;MsgBox ($MB_SYSTEMMODAL,"Debug","$sFile=""" & $sFile & """" & @CRLF & _
						;						"OuputFile=""" & $MyPath & $sOutputFolder & "\" & GetFileName($sFile) &".bmp""" & @CRLF & _
						;						"$sOutput=" & $sOutput & "<<")
					EndIf
					GUICtrlSetData($idProgressBar, Int($iCurrentFileImageNumber/$iFileImageNumber*100)); Update ProgressBar					
					;MsgBox ($MB_SYSTEMMODAL,"Debug","Stop !" )
				EndIf
			EndIf
			
		WEnd
		FileClose($s)
	EndIf
	Sleep(300)
	;MsgBox ($MB_SYSTEMMODAL,"Debug","Stop !")
EndFunc

Func FilenameExistsInBase()
	; >> Checks if the last filename in the database $aFilenameAfterConvert[] is redundant
	; Return is Boolean (True=Exists)
	Local $iOuputFileNumber, $bResult, $i
	$iOuputFileNumber=UBound($aFilenameAfterConvert)
	If $iOuputFileNumber>1 Then
		$bResult = False
		For $i=0 to $iOuputFileNumber-2 Step 1
			If $aFilenameAfterConvert[$iOuputFileNumber-1]=$aFilenameAfterConvert[$i] Then
				$bResult = True
				ExitLoop
			EndIf
		Next
	Else
		$bResult = False
	EndIf
	Return $bResult
EndFunc

Func CheckEthosNamingRules($sName)
	; >> Test if $sName comply with the third ETHOS naming rule 
	;   Rule #3 : At least one of these characters in the name : a~z or $sMandatoryChar=" ;[]+="
	;   Return the number of mandatory characters
	Local $iRules=0, $i, $cCurrentCharInName
	For $i=1 to StringLen($sName) Step 1
		$cCurrentCharInName=StringMid($sName,$i,1)
		If (Asc($cCurrentCharInName)>96 AND Asc($cCurrentCharInName)<123) Then $iRules+=1 ; a~z
		For $j=1 to StringLen($sMandatoryChar) Step 1
			If $cCurrentCharInName=StringMid($sMandatoryChar,$j,1) Then $iRules+=1 ; " ;[]+="
		Next
	Next
	Return $iRules
EndFunc

Func SetSuffix($iN,$sBase, $iLevel=1)
	; >> Generate a string based on $sBase string and $iN
    ; Example : SetSuffix(4,"abcde") => $sSuffix="d"
	;			SetSuffix(6,"abcde") => $sSuffix="aa"
	;			SetSuffix(7,"abcde") => $sSuffix="ab"
	;			SetSuffix(354,"abcde") => $sSuffix="bceb"
	Local $iMuliple, $iBaseSize, $iPlace
	If $iLevel=1 Then $sSuffix=""
    If StringLen($sSuffix)<$iLevel Then $sSuffix="*" & $sSuffix
    $iBaseSize=StringLen($sBase)
    $iMuliple=Int(($iN-1)/$iBaseSize)
 	$iPlace=$iN-$iMuliple*$iBaseSize
	$sSuffix=StringMid($sBase,$iPlace,1) & StringRight($sSuffix,StringLen($sSuffix)-1)
	If $iMuliple>0 Then SetSuffix($iMuliple,$sBase,$iLevel+1)
EndFunc


Func CountPictureFileToConvert()
	; >> Count number of file to convert
	$iFileImageNumber=0
	$s = FileFindFirstFile($MyPath & '*')
	If $s <> -1 Then
		While 1
			$sFile = FileFindNextFile($s)
			If @error Then ExitLoop
			If Not @extended Then
				If IsPictureFile($sFile) Then $iFileImageNumber +=1
			EndIf
		WEnd
		FileClose($s)
	EndIf
EndFunc

Func DoubleClick($hWnd, $iMsg, $wParam, $lParam)
	; >> Detect double click on main GUI to run dicovery of picture files and to run dicovery of Ethos SD Card drive
    Local $DbClick = BitAND($wParam, 0x00010000)
    If($DbClick) AND ($hWnd=$iGui) AND ($bConvertInProgress=False) Then
		CountPictureFileToConvert()
		FindEthosSD()
		SetPreferenceOnGUI()
        ;MsgBox($MB_SYSTEMMODAL, "Debug", "DoubleClick " & $iMsg & " / " & $idLabel1)
    EndIf
EndFunc

Func FindEthosSD()
	; >> Try to find a REMOVABLE drive with this content :
	; 	   \audio\en (directory)
	;	   \audio\fr (directory)
	;	   \bitmaps\user (directory)
	Local $aDrive = DriveGetDrive($DT_REMOVABLE)
	Local $bIsEthosContent, $isDirectory, $i, $j, $sTest=""
	Local $aTreeToTest[3]=["\audio\en","\audio\fr","\bitmaps\user"]
	If @error Then
		; An error occurred when retrieving the drives.
		$sEthosSDDrive=""
	Else
		For $i = 1 To $aDrive[0] Step 1
			$bIsEthosContent=True
			$sTest=$sTest & ">> Drive " & $i & "/" & $aDrive[0] & " >> " & $aDrive[$i] & @CRLF
			For $j=0 to UBound($aTreeToTest)-1 Step 1
				If Not FileExists($aDrive[$i] & $aTreeToTest[$j]) Then
					$bIsEthosContent=False
					$sTest=$sTest & "T" & ($j+1) & "/" & UBound($aTreeToTest) & " : " & $aDrive[$i] & $aTreeToTest[$j] & " > NotExist" 
				Else
					$sTest=$sTest & "T" & ($j+1) & "/" & UBound($aTreeToTest) & " : " & $aDrive[$i] & $aTreeToTest[$j] & " >  Exist  > " 
					$isDirectory = StringInStr(FileGetAttrib($aDrive[$i] & $aTreeToTest[$j]),"D")
					If Number($isDirectory)=0 Then
						$sTest=$sTest & "NotDirectory"
						$bIsEthosContent=False
					Else
						$sTest=$sTest & "Is Directory"
					EndIf
				EndIf
				$sTest=$sTest & @CRLF
			Next
			If $bIsEthosContent Then
				$sTest=$sTest & "Final : It's a valid ETHOS drive" & @CRLF
				$sEthosSDDrive=$aDrive[$i]
				ExitLoop
			Else
				$sTest=$sTest & "Final : It's not a valid ETHOS drive" & @CRLF
			EndIf
		Next
		;MsgBox($MB_SYSTEMMODAL, "Debug", $sTest)
	Endif
EndFunc

Func SetBool(ByRef $bBool)
	; >> Set boolean to True for Ternary conditionnal statement
	$bBool = True
EndFunc
Func UnsetBool(ByRef $bBool)
	; >> Set boolean to False for Ternary conditionnal statement
	$bBool = False
EndFunc

Func SetMessageOnGUI()
	; >> Update first message on main GUI
	Local $sMessage
	If $iFileImageNumber>0 Then
		$sMessage=$iFileImageNumber & " file(s) found to convert to the BMP format used by FrSky's ETHOS OS" & @CRLF
	Else
		$sMessage="No file to convert to the BMP format used by FrSky's ETHOS OS" & @CRLF
	EndIf
	If $sEthosSDDrive<>"" Then
		$sMessage=$sMessage & "The USB drive of the radio control SD card is detected (" & StringUpper($sEthosSDDrive) & " drive)" & @CRLF
	Else
		$sMessage=$sMessage & "No Ethos SD card" & @CRLF
	EndIf
	If $iFileImageNumber>0 Then
		If $bTransferSD=True And $sEthosSDDrive<>"" Then
			$sMessage=$sMessage & "► Ready to convert AND transfer to SD card !"
		EndIf
		If $bTransferSD=True And $sEthosSDDrive="" Then
			$sMessage=$sMessage & "► Ready to convert !"
		EndIf
	Else
		$sMessage=$sMessage & "Nothing to do in current directory"
	EndIf
	GUICtrlSetData($idLabel1,$sMessage)
EndFunc

Func SetPreferenceOnGUI()
	; >> Update GUI to match user preferences (checkbox, radiobutton, ...)
	Local $Value
	; on main Gui
	SetMessageOnGUI()
	($bAutorename=True) ? (GUICtrlSetState($idChkBox1, $GUI_CHECKED)) : (GUICtrlSetState($idChkBox1, $GUI_UNCHECKED))
	($bOverwriteComputer=True) ? (GUICtrlSetState($idChkBox2, $GUI_CHECKED)) : (GUICtrlSetState($idChkBox2, $GUI_UNCHECKED))
	($bTransferSD=True) ? (GUICtrlSetState($idChkBox3, $GUI_CHECKED)) : (GUICtrlSetState($idChkBox3, $GUI_UNCHECKED))
	If $iFileImageNumber>0 Then
		GUICtrlSetState($idConvert_Btn, $GUI_ENABLE)
		If $bTransferSD And $sEthosSDDrive<>"" Then ; Change text on Convert button
			GUICtrlSetData($idConvert_Btn,"Convert and transfer")
		Else
			GUICtrlSetData($idConvert_Btn,"Convert")
		Endif
	Else
		GUICtrlSetData($idConvert_Btn,"Nothing to do")
		GUICtrlSetState($idConvert_Btn, $GUI_DISABLE)
	EndIf
	; on preference Gui
	($bAvoidUppercase=True) ? (GUICtrlSetState($idGuiPrefUppercase, $GUI_CHECKED)) : (GUICtrlSetState($idGuiPrefUppercase, $GUI_UNCHECKED))
	($bLimitLength=True) ? (GUICtrlSetState($idGuiPrefCountChar, $GUI_CHECKED)) : (GUICtrlSetState($idGuiPrefCountChar, $GUI_UNCHECKED))
	GUICtrlSetData($idGuiPrefCharNumber,$iLimitLength)
	GUICtrlSetData($idGuiPrefCountChar,"Limit the name to " & $iLimitLength & " characters")
	($bLimitSpecialChar=True) ? (GUICtrlSetState($idGuiPrefSpecialChar, $GUI_CHECKED)) : (GUICtrlSetState($idGuiPrefSpecialChar, $GUI_UNCHECKED))
	GUICtrlSetData($idGuiCharSet,$sSpecialChar)
	If ($bOverwriteSD=True) Then
		GUICtrlSetState($idGuiPrefOverwrite, $GUI_CHECKED)
		GUICtrlSetState($idGuiPrefDontTransfer, $GUI_UNCHECKED)
	Else
		GUICtrlSetState($idGuiPrefOverwrite, $GUI_UNCHECKED)
		GUICtrlSetState($idGuiPrefDontTransfer, $GUI_CHECKED)
	Endif
EndFunc

Func LoadPreferences()
	; >> If INI file exist, load preferences
	; >> Otherwise set Default preferences
	Local $hFileOpen, $bSuccess, $sIniContent
	$bSuccess = True
	If FileExists($sIniFile) Then
		$hFileOpen = FileOpen($sIniFile, $FO_READ+$FO_UTF8)
		If $hFileOpen = -1 Then
			;MsgBox($MB_SYSTEMMODAL, "Debug", "An error occurred while loading preferences from INI file.")
			$bSuccess = False
		Else
			FileClose($hFileOpen) ; Close properly the INI file
			(IniRead($sIniFile, "AutoRenamePref", "Autorename", $bDAutorename)="True") ? (SetBool($bAutorename)) : (UnsetBool($bAutorename))
			;$bAvoidUppercase=IniRead($sIniFile, "AutoRenamePref", "AvoidUpperCase", $bDAvoidUppercase)
			(IniRead($sIniFile, "AutoRenamePref", "AvoidUpperCase", $bDAvoidUppercase)="True") ? (SetBool($bAvoidUppercase)) : (UnsetBool($bAvoidUppercase))
			(IniRead($sIniFile, "AutoRenamePref", "CheckNameLength", $bDLimitLength)="True") ? (SetBool($bLimitLength)) : (UnsetBool($bLimitLength))
			$iLimitLength=Number(IniRead($sIniFile, "AutoRenamePref", "MaxNameLength", $iDLimitLength))
			(IniRead($sIniFile, "AutoRenamePref", "LimitSpecialChar", $bDLimitSpecialChar)="True") ? (SetBool($bLimitSpecialChar)) : (UnsetBool($bLimitSpecialChar))
			$sSpecialChar=IniRead($sIniFile, "AutoRenamePref", "SpecialCharAllowed", $sDSpecialChar)
			(IniRead($sIniFile, "ConvertOnComputer", "OverwriteOnComputer", $bDOverwriteComputer)="True") ? (SetBool($bOverwriteComputer)) : (UnsetBool($bOverwriteComputer))
			(IniRead($sIniFile, "TransferToSDcard", "AutoTransferToSDcard", $bDTransferSD)="True") ? (SetBool($bTransferSD)) : (UnsetBool($bTransferSD))
			(IniRead($sIniFile, "TransferToSDcard", "OverwriteOnSDcard", $bDOverwriteSD)="True") ? (SetBool($bOverwriteSD)) : (UnsetBool($bOverwriteSD))
			;MsgBox($MB_SYSTEMMODAL, "DEBUG", "ContentFile" & @CRLF & "$bAutorename=" & $bAutorename & @CRLF & "$bAvoidUppercase=" & $bAvoidUppercase & @CRLF & _
			;		"$bLimitLength=" & $bLimitLength & @CRLF & "$iLimitLength=" & $iLimitLength & @CRLF & _
			;		"$bLimitSpecialChar=" & $bLimitSpecialChar & @CRLF & "$sSpecialChar=" & $sSpecialChar & @CRLF & _
			;		"$bOverwriteComputer=" & $bOverwriteComputer & @CRLF & "$bTransferSD=" & $bTransferSD & @CRLF & _
			;		"$bOverwriteSD=" & $bOverwriteSD)
			;DEBUG If ($bAvoidUppercase=True) Then MsgBox($MB_SYSTEMMODAL, "DEBUG", "$bAvoidUppercas= True") 
		EndIf
	Else
		;MsgBox($MB_SYSTEMMODAL, "DEBUG", "No ini file" & $sIniContent)
		$bAutorename=$bDAutorename
		$bOverwriteComputer=$bDOverwriteComputer
		$bTransferSD=$bDTransferSD
		$bAvoidUppercase=$bDAvoidUppercase
		$bLimitLength=$bDLimitLength
		$iLimitLength=$iDLimitLength
		$bLimitSpecialChar=$bDLimitSpecialChar
		$sSpecialChar=$sDSpecialChar
		$bOverwriteSD=$bDOverwriteSD
	EndIf	
	Return $bSuccess
EndFunc

Func SavePreferences()
	; >> Test if Preferences are set to Default
	; >>   Default    -> No INI file created and INI file deletion if exists
	; >>   Other case -> Create INI file
	Local $hFileOpen, $bSuccess, $bCreateINI
	$bCreateINI = True
	If (($bAutorename=$bDAutorename) AND ($bOverwriteComputer=$bDOverwriteComputer) AND ($bTransferSD=$bDTransferSD) AND _
		($bAvoidUppercase=$bDAvoidUppercase) AND ($bLimitLength=$bDLimitLength) AND ($iLimitLength=$iDLimitLength) AND _
		($bLimitSpecialChar=$bDLimitSpecialChar) AND ($sSpecialChar=$sDSpecialChar) AND ($bOverwriteSD=$bDOverwriteSD)) _
		Then $bCreateINI = False	
	If FileExists($sIniFile) Then
		$bSuccess = True
		If FileDelete($sIniFile)=0 Then $bSuccess = False
	Endif 
	If $bCreateINI = True Then ; Need to create INI file because it's not Default preferences
		$hFileOpen = FileOpen($sIniFile, $FO_APPEND+$FO_UTF8)
		If $hFileOpen = -1 Then
			MsgBox($MB_SYSTEMMODAL, "", "An error occurred while writing preferences on INI file.")
			$bSuccess = False
		Else
			FileClose($hFileOpen) ; An empty INI file is created
			IniWrite($sIniFile, "AutoRenamePref", "Autorename", $bAutorename)
			IniWrite($sIniFile, "AutoRenamePref", "AvoidUpperCase", $bAvoidUppercase)
			IniWrite($sIniFile, "AutoRenamePref", "CheckNameLength", $bLimitLength)
			IniWrite($sIniFile, "AutoRenamePref", "MaxNameLength", $iLimitLength)
			IniWrite($sIniFile, "AutoRenamePref", "LimitSpecialChar", $bLimitSpecialChar)
			IniWrite($sIniFile, "AutoRenamePref", "SpecialCharAllowed", $sSpecialChar)
			IniWrite($sIniFile, "ConvertOnComputer", "OverwriteOnComputer", $bOverwriteComputer)
			IniWrite($sIniFile, "TransferToSDcard", "AutoTransferToSDcard", $bTransferSD)
			IniWrite($sIniFile, "TransferToSDcard", "OverwriteOnSDcard", $bOverwriteSD)
			$bSuccess = True
		EndIf
	EndIf
	Return $bSuccess
EndFunc

Func CleanSpecialChar($sInput)
	; >> Clean the user charset to comply with $sDSpecialCharSet="é#([-è_çà@)]=+ù.!<>,.;!§"
	Local $YLoop, $XLoop, $sCharIn, $sCharOut, $sCharSet
	$sCharOut = ""
	For $YLoop = 1 to StringLen ($sInput) Step 1
		$sCharIn = StringMid ($sInput, $YLoop, 1)
		For $XLoop = 1 to StringLen ($sDSpecialCharSet) Step 1
			$sCharSet = StringMid ($sDSpecialCharSet, $XLoop, 1)
			If $sCharIn=$sCharSet Then
				$sCharOut=$sCharOut & $sCharSet
				ExitLoop
			Endif
		Next
	Next
	Return $sCharOut
EndFunc

Func GuiMain()
	; >> To design the main Gui
	$iGui = GUICreate("Convert to ETHOS BMP file format", $iGUIWidth, $iGUIHeight)
	$aParentWin_Pos = WinGetPos($iGui, "")
	; Label to display info before converting
	$idLabel1 = GUICtrlCreateLabel("3 files found to convert to the BMP format used by FrSky's ETHOS OS" & @CRLF & "The USB drive of the radio control SD card is detected (F:/ drive)" & @CRLF & "► Ready to convert !", $Label1Margin, $Label1Margin,$iGUIWidth-2*$Label1Margin,$Label1Height,$SS_CENTER,$WS_EX_STATICEDGE)
	;GUICtrlSetFont ($idLabel1,9, $FW_BOLD)
	GUICtrlSetFont ($idLabel1,10);
	
	; CheckBox for AUTORENAME
	$idChkBox1= GUICtrlCreateCheckbox("Automatically rename the output files if necessary", 20, $Label1Height+$SpaceLb1ChB1)
	GUICtrlSetTip($idChkBox1, "To control auto rename click on ""Preferences""")
	GUICtrlSetState($idChkBox1, $GUI_CHECKED)
	; CheckBox for OVERWRITE
	$idChkBox2= GUICtrlCreateCheckbox("Overwrite if output files exist on this computer", 20, $Label1Height+$SpaceLb1ChB1+$SpaceBtwChB)
	GUICtrlSetTip($idChkBox2, "Checks if the files exist in the FrSkyEthosBMP directory")
	GUICtrlSetState($idChkBox2, $GUI_CHECKED)
	; CheckBox for COPY TO SD CARD
	$idChkBox3= GUICtrlCreateCheckbox("Transfer the BMP files to the SD card of the radio control", 20, $Label1Height+$SpaceLb1ChB1+2*$SpaceBtwChB)
	GUICtrlSetTip($idChkBox3, "Only works if the USB drive of the SD card of the radio control is detected")
	GUICtrlSetState($idChkBox3, $GUI_CHECKED)
	;GUICtrlSetState($idChkBox3, $GUI_DISABLE)

	; Create an "CONVERT" button
	$idConvert_Btn = GUICtrlCreateButton("Convert and transfer", $iGUIWidth-160, $iGUIHeight-$Label1Margin-25, 150, 25, $BS_DEFPUSHBUTTON)
	;GUICtrlSetState($idConvert_Btn, $GUI_FOCUS)
	; Create a "CANCEL" button
	$idCancel_Btn = GUICtrlCreateButton("Cancel", $iGUIWidth-250, $iGUIHeight-$Label1Margin-25, 80, 25)
	; Create a "PREFERENCES" button
	$idPreferences_Btn = GUICtrlCreateButton("Preferences", $Label1Margin, $iGUIHeight-$Label1Margin-25, 80, 25)
EndFunc

Func GuiTransfer()
	; >> Redesign Guimain 
	; Disable all control
	GUICtrlSetState($idConvert_Btn, $GUI_DISABLE)
	GUICtrlSetState($idCancel_Btn, $GUI_DISABLE)
	GUICtrlSetState($idPreferences_Btn, $GUI_DISABLE)
	GUICtrlSetState($idChkBox1, $GUI_DISABLE)
	GUICtrlSetState($idChkBox2, $GUI_DISABLE)
	GUICtrlSetState($idChkBox3, $GUI_DISABLE)
	GUICtrlDelete($idLabel1)
	; Progress bar with label
	$idProgressBar = GUICtrlCreateProgress($Label1Margin, $Label1Margin,$iGUIWidth-2*$Label1Margin,30)
	GUICtrlSetTip($idProgressBar, "Init convert")
	GUICtrlSetData($idProgressBar, 0)
	$idLabel1 = GUICtrlCreateLabel("", $Label1Margin, 50,$iGUIWidth-2*$Label1Margin,20,$SS_CENTER)
EndFunc

Func GuiPreference()
	; >> To design the Preference Gui
	$idGuiPref = GUICreate("Convert to BMP >> Preferences", $iGuiPrefWidth, $iGuiPrefHeight, $aParentWin_Pos[0] + 100, $aParentWin_Pos[1] + 100, -1, -1, $iGui)
	; Icon
	$idGuiPrefIcon=GUICtrlCreateButton("", $iGuiPrefWidth/2-25, 3, 38, 38, $BS_ICON)
    GUICtrlSetImage(-1, @ScriptName, 0, 1)
	; Version information
	$idGuiPrefLabelInfo=GUICtrlCreateLabel("GNU GPL2 - " & $sVersion & "  •  by Ceeb182  •  https://github.com/Ceeb182", $Label1Margin, $Label1Margin+$iIcon,$iGuiPrefWidth-2*$Label1Margin,20,$SS_CENTER)
	; Group AUTO-RENAME
	$idGuiPrefGroupRename = GUICtrlCreateGroup("Auto rename", $Label1Margin, $Label1Margin+$iIcon+20,$iGuiPrefWidth-2*$Label1Margin ,100)
	$iGuiPregYPos=$Label1Margin+$iIcon+20+100+5
	GUICtrlSetFont ($idGuiPrefGroupRename,9, $FW_BOLD)
	$idGuiPrefUppercase = GUICtrlCreateCheckbox("Comply with ETHOS naming rules", $Label1Margin+20, $iIcon+50)
	GUICtrlSetTip(-1, "Forbid accented characters like é, è, ç, ù" & @CRLF & _
					  "Allowed characters : A~Z, a~z, 0~9, ()!-_@#;[]+=`Space`" & @CRLF & _
					  "At least one of these characters in the name : a~z or ;[]+=`Space`")
	$idGuiPrefCountChar = GUICtrlCreateCheckbox("Limit the name to 11 characters", $Label1Margin+20, $iIcon+50+$SpaceBtwChB)
	$idGuiPrefCharNumber = GUICtrlCreateInput("11", $Label1Margin+215, $iIcon+50+$SpaceBtwChB, 50, 20)
	GUICtrlCreateUpdown($idGuiPrefCharNumber,$UDS_HORZ)
	GUICtrlSetPos($idGuiPrefCharNumber, $Label1Margin+215, $iIcon+50+$SpaceBtwChB, 50, 20)
	GUICtrlSetLimit($idGuiPrefCharNumber, 2,1)
	$idGuiPrefSpecialChar = GUICtrlCreateCheckbox("Allow only this special character set :", $Label1Margin+20, $iIcon+50+2*$SpaceBtwChB)
	GUICtrlSetTip(-1,"Only `Space` ;[]+=()!-_@# are allowed")
	$idGuiCharSet = GUICtrlCreateInput("-_+!", $Label1Margin+215, $iIcon+50+2*$SpaceBtwChB, 220, 20)
	GUICtrlSetTip(-1,"Only `Space` ;[]+=()!-_@# are allowed")
	GUICtrlCreateGroup("", -99, -99, 1, 1) ;close group
	; Group TRANSFER TO SD CARD
	$idGuiPrefGroupTransfer = GUICtrlCreateGroup("Tranfer to radio SD Card", $Label1Margin, $iGuiPregYPos,$iGuiPrefWidth-2*$Label1Margin ,100)
	GUICtrlSetFont ($idGuiPrefGroupTransfer,9, $FW_BOLD);
	$idGuiLabelTransfer = GUICtrlCreateLabel("If the file already exists on radio SD Card :", $Label1Margin+20, $iGuiPregYPos+20)
	$idGuiPrefOverwrite = GUICtrlCreateRadio("Overwrite this file",$Label1Margin+20,$iGuiPregYPos+20+$SpaceBtwChB)
	$idGuiPrefDontTransfer = GUICtrlCreateRadio("Do not transfer this file",$Label1Margin+20,$iGuiPregYPos+20+2*$SpaceBtwChB)
	GUICtrlSetState($idGuiPrefOverwrite, $GUI_CHECKED)	
	GUICtrlCreateGroup("", -99, -99, 1, 1) ;close group
	; Buttons
	$idGuiPrefOk = GUICtrlCreateButton("Ok", $iGuiPrefWidth-160, $iGuiPrefHeight-$Label1Margin-25, 150, 25, $BS_DEFPUSHBUTTON)
	$idGuiPrefDefault = GUICtrlCreateButton("Default settings", $Label1Margin, $iGuiPrefHeight-$Label1Margin-25, 150, 25)
EndFunc

Func IsPictureFile($file)
	; >> Check for picture files which match with $ImageExt[] extension
	Local $ExtNumber, $FileExt, $bResult
	$ExtNumber=UBound($ImageExt)
	$FileExt=GetFileExt($file)
	$bResult=False
	For $YLoop = 0 to $ExtNumber-1 Step 1
		If $ImageExt[$YLoop]=$FileExt Then
			$bResult=True
			ExitLoop
		EndIf
	Next
	Return $bResult
EndFunc

Func GetFileExt($file)
	; >> Get file extension
	Local $YLoop, $fl_Ext
	$fl_Ext=""
	For $YLoop = StringLen ($file) to 1 Step -1
		If StringMid ($file, $YLoop, 1) == "." Then
			$fl_Ext = StringMid ($file, $YLoop)
			ExitLoop
		EndIf
	Next
	Return $fl_Ext
EndFunc

Func GetFileName($file)
	; >> Get file name
	Local $YLoop, $fl_Name
	$fl_Ext=""
	For $YLoop = StringLen ($file) to 1 Step -1
		If StringMid ($file, $YLoop, 1) == "." Then
			$fl_Name = StringLeft ($file, $YLoop-1)
			ExitLoop
		EndIf
	Next
	Return $fl_Name
EndFunc

Func MyErrFunc()
	; >> Error management
	$HexNumber=hex($oMyError.number,8)
	Msgbox(0,"COM Error Test","We intercepted a COM Error !"     & @CRLF & @CRLF & _
             "err.description is: " & @TAB & $oMyError.description  & @CRLF & _
             "err.windescription:"   & @TAB & $oMyError.windescription & @CRLF & _
             "err.number is: "       & @TAB & $HexNumber             & @CRLF & _
             "err.lastdllerror is: " & @TAB & $oMyError.lastdllerror & @CRLF & _
             "err.scriptline is: "   & @TAB & $oMyError.scriptline   & @CRLF & _
             "err.source is: "       & @TAB & $oMyError.source       & @CRLF & _
             "err.helpfile is: "     & @TAB & $oMyError.helpfile     & @CRLF & _
             "err.helpcontext is: " & @TAB & $oMyError.helpcontext _
            )
	SetError(1) ; to check for after this function returns
Endfunc