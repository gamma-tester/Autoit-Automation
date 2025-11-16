;==============================================================
; Photo Copy & Resize – AutoIt3 version
;==============================================================
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <File.au3>
#include <GDIPlus.au3>

Global $hGui, $srcInput, $destInput, $nameInput, $statusLabel, $srcBrowseBtn, $destBrowseBtn, $processBtn, $progressBar, $previewPic, $recentBtn
Global $sScriptDir = StringRegExpReplace(@ScriptDir, "\\$", "") ; strip trailing back-slash
Global $sIrfanViewPath = ""
Global $aRecentFiles[0]
Global $hPreviewBitmap = 0

;-------------------------------------------------------------- GUI
$hGui = GUICreate("Photo Copy & Resize", 800, 400, -1, -1, -1, $WS_EX_TOPMOST)
GUISetFont(9, 400, 0, "Segoe UI")

; Left side - Controls
GUICtrlCreateLabel("Source Image File:", 20, 20, 350, 20)
$srcInput = GUICtrlCreateInput("", 20, 40, 350, 24)
GUICtrlSetState(-1, $GUI_DISABLE)
$srcBrowseBtn = GUICtrlCreateButton("Browse", 380, 40, 80, 24)

GUICtrlCreateLabel("Destination Folder:", 20, 80, 350, 20)
$destInput = GUICtrlCreateInput("", 20, 100, 350, 24)
$destBrowseBtn = GUICtrlCreateButton("Browse", 380, 100, 80, 24)

GUICtrlCreateLabel("Client Number:", 20, 140, 350, 20)
$nameInput = GUICtrlCreateInput("", 20, 160, 80, 24)
GUICtrlSetLimit(-1, 5)
GUICtrlSetStyle(-1, $ES_NUMBER) ; Only allow numbers

$processBtn = GUICtrlCreateButton("Copy & Resize", 20, 200, 120, 30)
$recentBtn = GUICtrlCreateButton("Recent Files", 150, 200, 120, 30)
$progressBar = GUICtrlCreateProgress(20, 240, 440, 20)
GUICtrlSetState($progressBar, $GUI_HIDE)

; Right side - Preview
GUICtrlCreateLabel("Image Preview:", 480, 20, 300, 20)
$previewPic = GUICtrlCreatePic("", 480, 40, 300, 300)
GUICtrlSetState($previewPic, $GUI_HIDE)

$statusLabel = GUICtrlCreateLabel("", 480, 350, 300, 40)
GUICtrlSetColor(-1, 0x008000)

;-------------------------------------------------------------- Load config
_LoadConfig()
_LoadIrfanViewPath()
_GDIPlus_Startup()
GUISetState()

;-------------------------------------------------------------- Message loop
While 1
    Switch GUIGetMsg()
        Case $GUI_EVENT_CLOSE
            ExitLoop
        Case $srcBrowseBtn
            _BrowseFile()
        Case $destBrowseBtn
            _BrowseFolder()
        Case $processBtn
            _RunBatch()
        Case $recentBtn
            _ShowRecentFiles()
        Case $srcInput
            ; Update preview when source file changes
            _UpdatePreview()
    EndSwitch
WEnd

; Cleanup
_GDIPlus_Shutdown()
Exit

;==============================================================
; Functions
;==============================================================
Func _BrowseFile()
    Local $sFile = FileOpenDialog("Select an image", @ScriptDir, "Images (*.jpg;*.jpeg;*.png)", 1)
    If @error Then Return
    
    ; Validate file type
    Local $sExt = StringLower(StringTrimLeft($sFile, StringInStr($sFile, ".", 0, -1)))
    If $sExt <> "jpg" And $sExt <> "jpeg" And $sExt <> "png" Then
        MsgBox(16, "Invalid File", "Please select a valid image file (JPG, JPEG, or PNG).")
        Return
    EndIf
    
    GUICtrlSetData($srcInput, $sFile)
    ; Force update preview immediately
    _UpdatePreview()
EndFunc

Func _BrowseFolder()
    Local $sFolder = FileSelectFolder("Select destination folder", "")
    If @error Then Return
    GUICtrlSetData($destInput, $sFolder)
EndFunc

Func _RunBatch()
    Local $sSrc  = StringStripWS(GUICtrlRead($srcInput), 3)
    Local $sDest = StringStripWS(GUICtrlRead($destInput), 3)
    Local $sName = StringStripWS(GUICtrlRead($nameInput), 3)

    GUICtrlSetData($statusLabel, "")

    If $sSrc = "" Or $sDest = "" Or $sName = "" Then
        _ShowStatus("Please fill in all fields.", 0)
        Return
    EndIf
    If StringLen($sName) <> 5 Or Not StringIsDigit($sName) Then
        _ShowStatus("Name must be exactly 5 digits.", 0)
        Return
    EndIf

    If StringRight($sDest, 1) = "\" Then $sDest = StringTrimRight($sDest, 1)

    ; Convert to absolute paths
    $sSrc = _GetAbsolutePath($sSrc)
    $sDest = _GetAbsolutePath($sDest)
    
    If Not FileExists($sSrc) Then
        _ShowStatus("Source file not found: " & $sSrc, 0)
        Return
    EndIf
    If Not FileExists($sDest) Then
        _ShowStatus("Destination folder not found: " & $sDest, 0)
        Return
    EndIf

    ; Process the image directly in AutoIt3
    _ProcessImage($sSrc, $sDest, $sName)
EndFunc

Func _LoadIrfanViewPath()
    Local $sConfig = $sScriptDir & "\config.file"
    If Not FileExists($sConfig) Then Return
    
    Local $aLines = FileReadToArray($sConfig)
    For $i = 0 To UBound($aLines) - 1
        Local $sLine = StringStripWS($aLines[$i], 3)
        If StringInStr($sLine, "IrfanViewPath") Then
            Local $sPath = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
            $sIrfanViewPath = $sPath
            ExitLoop
        EndIf
    Next
EndFunc

Func _GetAbsolutePath($sPath)
    ; Convert relative path to absolute path
    If StringLeft($sPath, 1) = "." Then
        ; Relative path starting with .\
        Return $sScriptDir & "\" & StringTrimLeft($sPath, 2)
    ElseIf Not StringInStr($sPath, ":") And StringLeft($sPath, 2) <> "\\" Then
        ; Relative path without drive letter or UNC
        Return $sScriptDir & "\" & $sPath
    Else
        ; Already absolute path
        Return $sPath
    EndIf
EndFunc

Func _ProcessImage($sSrc, $sDest, $sName)
    ; Build output filename
    Local $sOutFile = $sDest & "\" & $sName & ".jpg"
    
    ; Check if output file already exists
    If FileExists($sOutFile) Then
        ; Temporarily remove topmost style to show dialog in front
        WinSetOnTop($hGui, "", 0)
        
        ; Ask user for overwrite confirmation (modal to main window)
        Local $iResponse = MsgBox(4 + 32 + 262144, "File Exists", "File " & $sName & ".jpg already exists." & @CRLF & "Do you want to overwrite it?")
        
        ; Restore topmost style
        WinSetOnTop($hGui, "", 1)
        
        If $iResponse <> 6 Then ; 6 = Yes
            _ShowStatus("Operation cancelled by user.", 0)
            Return
        EndIf
    EndIf
    
    ; Get IrfanView path from config
    If $sIrfanViewPath = "" Then
        _LoadIrfanViewPath()
    EndIf
    
    ; Check if IrfanView is available
    If Not FileExists($sIrfanViewPath) Then
        _ShowStatus("IrfanView not found at: " & $sIrfanViewPath, 0)
        Return
    EndIf
    
    ; Build IrfanView command for resize and conversion
    Local $sCmd = '"' & $sIrfanViewPath & '" "' & $sSrc & '" /resize_long=800 /aspectratio /resample /convert="' & $sOutFile & '" /jpgq=85'
    
    ; Show progress bar
    GUICtrlSetState($progressBar, $GUI_SHOW)
    GUICtrlSetData($progressBar, 0)
    _ShowStatus("Starting image processing...", 1)
    
    ; Run IrfanView
    Local $iPID = Run($sCmd, "", @SW_HIDE)
    If @error Then
        GUICtrlSetState($progressBar, $GUI_HIDE)
        _ShowStatus("Failed to start IrfanView.", 0)
        Return
    EndIf
    
    ; Simulate progress for longer operations
    For $i = 10 To 90 Step 20
        GUICtrlSetData($progressBar, $i)
        Sleep(100)
        If Not ProcessExists($iPID) Then ExitLoop
    Next
    
    ; Wait for process to complete
    ProcessWaitClose($iPID)
    Local $iExitCode = @extended
    
    ; Complete progress
    GUICtrlSetData($progressBar, 100)
    Sleep(500)
    GUICtrlSetState($progressBar, $GUI_HIDE)
    
    ; Check if output file was created
    If FileExists($sOutFile) Then
        _ShowStatus("Image processed successfully: " & $sName & ".jpg", 1)
        ; Add to recent files
        _AddToRecentFiles($sSrc)
    Else
        _ShowStatus("Failed to process image. Check if IrfanView is working properly.", 0)
    EndIf
EndFunc

Func _UpdatePreview()
    Local $sSrc = StringStripWS(GUICtrlRead($srcInput), 3)
    If $sSrc = "" Or Not FileExists($sSrc) Then
        _ClearPreview()
        Return
    EndIf
    
    ; Validate file type
    Local $sExt = StringLower(StringTrimLeft($sSrc, StringInStr($sSrc, ".", 0, -1)))
    If $sExt <> "jpg" And $sExt <> "jpeg" And $sExt <> "png" Then
        _ClearPreview()
        Return
    EndIf
    
    ; Clear previous preview
    _ClearPreview()
    
    ; Load image using GDI+
    Local $hImage = _GDIPlus_ImageLoadFromFile($sSrc)
    If @error Then
        _ClearPreview()
        Return
    EndIf
    
    ; Get image dimensions
    Local $iWidth = _GDIPlus_ImageGetWidth($hImage)
    Local $iHeight = _GDIPlus_ImageGetHeight($hImage)
    
    ; Calculate scaled dimensions to fit 300x300 preview
    Local $iNewWidth, $iNewHeight
    If $iWidth > $iHeight Then
        $iNewWidth = 300
        $iNewHeight = Int($iHeight * 300 / $iWidth)
    Else
        $iNewHeight = 300
        $iNewWidth = Int($iWidth * 300 / $iHeight)
    EndIf
    
    ; Create bitmap for preview
    Local $hBitmap = _GDIPlus_BitmapCreateFromScan0($iNewWidth, $iNewHeight)
    Local $hGraphics = _GDIPlus_ImageGetGraphicsContext($hBitmap)
    _GDIPlus_GraphicsSetInterpolationMode($hGraphics, $GDIP_INTERPOLATIONMODE_HIGHQUALITYBICUBIC)
    _GDIPlus_GraphicsDrawImageRect($hGraphics, $hImage, 0, 0, $iNewWidth, $iNewHeight)
    
    ; Save bitmap to file for Pic control
    Local $sTempBMP = @TempDir & "\preview_" & @MSEC & ".bmp"
    _GDIPlus_ImageSaveToFile($hBitmap, $sTempBMP)
    
    ; Cleanup GDI+ objects
    _GDIPlus_GraphicsDispose($hGraphics)
    _GDIPlus_ImageDispose($hBitmap)
    _GDIPlus_ImageDispose($hImage)
    
    ; Show preview
    If FileExists($sTempBMP) Then
        GUICtrlSetImage($previewPic, $sTempBMP)
        GUICtrlSetState($previewPic, $GUI_SHOW)
        $hPreviewBitmap = $hBitmap
    Else
        _ClearPreview()
    EndIf
EndFunc

Func _ClearPreview()
    GUICtrlSetState($previewPic, $GUI_HIDE)
    GUICtrlSetImage($previewPic, "")
    
    ; Clean up temporary preview files
    Local $aFiles = _FileListToArray(@TempDir, "preview_*.bmp", 1)
    If Not @error Then
        For $i = 1 To $aFiles[0]
            FileDelete(@TempDir & "\" & $aFiles[$i])
        Next
    EndIf
EndFunc

Func _ShowStatus($sText, $bOK)
    GUICtrlSetData($statusLabel, $sText)
    If $bOK Then
        GUICtrlSetColor($statusLabel, 0x008000)
    Else
        GUICtrlSetColor($statusLabel, 0xFF0000)
    EndIf
EndFunc

Func _AddToRecentFiles($sFile)
    ; Add file to recent files array (max 5 files)
    ; Use simple array manipulation instead of _Array functions
    Local $iSize = UBound($aRecentFiles)
    
    ; Create new array with the new file at the beginning
    Local $aNewArray[$iSize + 1]
    $aNewArray[0] = $sFile
    
    ; Copy existing files
    For $i = 0 To $iSize - 1
        $aNewArray[$i + 1] = $aRecentFiles[$i]
    Next
    
    ; Keep only first 5 files
    If UBound($aNewArray) > 5 Then
        Local $aTemp[5]
        For $i = 0 To 4
            $aTemp[$i] = $aNewArray[$i]
        Next
        $aRecentFiles = $aTemp
    Else
        $aRecentFiles = $aNewArray
    EndIf
EndFunc

Func _ShowRecentFiles()
    Local $iSize = UBound($aRecentFiles)
    If $iSize = 0 Then
        ; Temporarily remove topmost style to show dialog in front
        WinSetOnTop($hGui, "", 0)
        MsgBox(64, "Recent Files", "No recent files processed yet.")
        ; Restore topmost style
        WinSetOnTop($hGui, "", 1)
        Return
    EndIf
    
    Local $sRecentList = "Recent Files:" & @CRLF & @CRLF
    For $i = 0 To $iSize - 1
        $sRecentList &= ($i + 1) & ". " & $aRecentFiles[$i] & @CRLF
    Next
    
    ; Temporarily remove topmost style to show dialog in front
    WinSetOnTop($hGui, "", 0)
    MsgBox(64, "Recent Files", $sRecentList)
    ; Restore topmost style
    WinSetOnTop($hGui, "", 1)
EndFunc

Func _LoadConfig()
    Local $sConfig = $sScriptDir & "\config.file"
    If Not FileExists($sConfig) Then Return
    Local $aLines = FileReadToArray($sConfig)
    For $i = 0 To UBound($aLines) - 1
        Local $sLine = StringStripWS($aLines[$i], 3)
        If StringInStr($sLine, "Folder-1") Then
            Local $sPath = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
            GUICtrlSetData($destInput, $sPath)
            ExitLoop
        EndIf
    Next
EndFunc
