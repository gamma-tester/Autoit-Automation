#Include-once
#include <GDIPlus.au3>
#include <INet.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <File.au3>

; === Microsoft Face API Configuration ===
Global $sSubscriptionKey = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
Global $sEndpoint       = "https://your-end-poing/"
Global $sOutputDir      = @ScriptDir & "\output"

; Global variables for preview
Global $hPreviewGUI, $hPreviewImage, $aDetectedFaces, $sCurrentImagePath, $sCurrentOutputPath
Global $iSelectedFace = -1  ; -1 means no face selected
Global $aFaceRegions[0]  ; Array to store face regions for click detection

; === Entry point ===
Main()

Func Main()
    ; Initialize GDI+ at the start
    _GDIPlus_Startup()
    
    ; ---- pick image ----
    Local $sImagePath = FileOpenDialog("Select Image for AI Face Detection", @ScriptDir, _
                                      "Images (*.jpg;*.jpeg;*.png;*.bmp)", 1)
    If @error Then
        MsgBox(0, "Info", "No image selected. Exiting.")
        _GDIPlus_Shutdown()
        Return
    EndIf

    ; ---- output folder ----
    If Not FileExists($sOutputDir) Then DirCreate($sOutputDir)

    ; ---- build output filename ----
    Local $sFileName = StringRegExpReplace($sImagePath, "^.*\\", "")
    $sFileName = StringRegExpReplace($sFileName, "\.[^\.]+$", "")
    Local $sOutputPath = $sOutputDir & "\" & $sFileName & "_cropped_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".jpg"

    ConsoleWrite("Starting Microsoft Face API detection..." & @CRLF)
    ConsoleWrite("Using Endpoint: " & $sEndpoint & @CRLF)

    ; ---- call API ----
    Local $aFaces = DetectFacesWithMicrosoft($sImagePath)

    If @error Then
        Local $iError = @error
        ConsoleWrite("Face API Error: " & $iError & @CRLF)

        If $iError = 1 Then
            MsgBox(16, "API Error", "Invalid API Key or Endpoint." & @CRLF & "Please check your Azure credentials.")
        ElseIf $iError = 2 Then
            MsgBox(16, "Network Error", "Could not connect to Microsoft Face API." & @CRLF & "Using center crop instead.")
        ElseIf $iError = 3 Then
            MsgBox(16, "API Feature Error", "Face API feature not available." & @CRLF & "Using center crop instead.")
        Else
            MsgBox(16, "Error", "Failed to call Face API. Error code: " & $iError)
        EndIf

        ; ---- fallback to center crop ----
        If CenterCrop($sImagePath, $sOutputPath) Then
            MsgBox(64, "Center Crop", "Image cropped to 1:1 ratio!" & @CRLF & "Saved: " & $sOutputPath)
            ShellExecute($sOutputDir)
        Else
            MsgBox(16, "Error", "Failed to crop image")
        EndIf
        _GDIPlus_Shutdown()
        Return
    EndIf

    ; ---- no faces found ----
    If UBound($aFaces) = 0 Then
        MsgBox(48, "No Faces", "Microsoft AI did not detect any faces in the image." & @CRLF & "Using center crop instead.")
        If CenterCrop($sImagePath, $sOutputPath) Then
            MsgBox(64, "Center Crop", "Image cropped to 1:1 ratio!" & @CRLF & "Saved: " & $sOutputPath)
            ShellExecute($sOutputDir)
        EndIf
        _GDIPlus_Shutdown()
        Return
    EndIf

    ; Store global variables for preview
    $aDetectedFaces = $aFaces
    $sCurrentImagePath = $sImagePath
    $sCurrentOutputPath = $sOutputPath

    ; ---- Show interactive preview with face selection ----
    ShowInteractivePreview($sImagePath, $aFaces)

    _GDIPlus_Shutdown()
EndFunc   ;==>Main

; ------------------------------------------------------------------
;  Show interactive preview with clickable face selection
; ------------------------------------------------------------------
Func ShowInteractivePreview($sImagePath, $aFaces)
    ; Load original image
    Local $hOriginalImage = _GDIPlus_ImageLoadFromFile($sImagePath)
    Local $iImgWidth = _GDIPlus_ImageGetWidth($hOriginalImage)
    Local $iImgHeight = _GDIPlus_ImageGetHeight($hOriginalImage)
    
    ; Calculate preview size (max 800x600 while maintaining aspect ratio)
    Local $iPreviewWidth, $iPreviewHeight
    Local $fAspectRatio = $iImgWidth / $iImgHeight
    
    If $iImgWidth > 800 Or $iImgHeight > 600 Then
        If $fAspectRatio > (800/600) Then
            $iPreviewWidth = 800
            $iPreviewHeight = 800 / $fAspectRatio
        Else
            $iPreviewHeight = 600
            $iPreviewWidth = 600 * $fAspectRatio
        EndIf
    Else
        $iPreviewWidth = $iImgWidth
        $iPreviewHeight = $iImgHeight
    EndIf
    
    ; Calculate scale factors
    Local $fScaleX = $iPreviewWidth / $iImgWidth
    Local $fScaleY = $iPreviewHeight / $iImgHeight
    
    ; Create preview GUI
    $hPreviewGUI = GUICreate("Face Detection - Click Face to Select", $iPreviewWidth + 20, $iPreviewHeight + 120)
    GUISetBkColor(0xF0F0F0)
    
    ; Create image control (make it clickable)
    $hPreviewImage = GUICtrlCreatePic("", 10, 10, $iPreviewWidth, $iPreviewHeight)
    GUICtrlSetCursor(-1, 2) ; Set cursor to hand when hovering over image
    
    ; Create buttons
    Local $hCropButton = GUICtrlCreateButton("Crop Selected Face", 10, $iPreviewHeight + 20, 150, 30)
    Local $hCancelButton = GUICtrlCreateButton("Cancel", 170, $iPreviewHeight + 20, 80, 30)
    
    ; Create status label
    Local $hStatusLabel = GUICtrlCreateLabel("Faces detected: " & UBound($aFaces) & " - Click the Face you want to crop", 10, $iPreviewHeight + 60, $iPreviewWidth, 20)
    Local $hSelectionLabel = GUICtrlCreateLabel("No face selected", 10, $iPreviewHeight + 85, $iPreviewWidth, 20)
    
    ; Create preview image with face rectangles and store face regions
    CreateInteractivePreviewImage($sImagePath, $aFaces, $iPreviewWidth, $iPreviewHeight, $fScaleX, $fScaleY)
    
    GUISetState(@SW_SHOW)
    
    ; Event loop
    While 1
        Local $nMsg = GUIGetMsg()
        Switch $nMsg
            Case $GUI_EVENT_CLOSE, $hCancelButton
                GUIDelete($hPreviewGUI)
                ExitLoop
            Case $hCropButton
                If $iSelectedFace = -1 Then
                    MsgBox(48, "No Face Selected", "Please click on a face to select it first.")
                Else
                    ; Crop the selected face
                    If CropDetectedFace($sCurrentImagePath, $sCurrentOutputPath, $aDetectedFaces[$iSelectedFace]) Then
                        Local $sMsg = "AI Face Detection Successful!" & @CRLF & _
                                     "Faces detected: " & UBound($aDetectedFaces) & @CRLF & _
                                     "Cropped selected face to 1:1 ratio" & @CRLF & _
                                     "Saved: " & $sCurrentOutputPath

                        If UBound($aDetectedFaces) > 1 Then
                            $sMsg &= @CRLF & @CRLF & "Note: " & UBound($aDetectedFaces) - 1 & " additional face(s) were detected but not cropped."
                        EndIf
                        MsgBox(64, "Success", $sMsg)
                        ShellExecute($sOutputDir)
                    Else
                        MsgBox(16, "Error", "Failed to crop selected face")
                    EndIf
                    GUIDelete($hPreviewGUI)
                    ExitLoop
                EndIf
            Case $hPreviewImage
                ; Handle click on image
                Local $aMousePos = GUIGetCursorInfo($hPreviewGUI)
                If IsArray($aMousePos) Then
                    Local $iMouseX = $aMousePos[0] - 10  ; Adjust for image position
                    Local $iMouseY = $aMousePos[1] - 10
                    
                    ; Check if click is within any face region
                    ConsoleWrite("Mouse Click: X=" & $iMouseX & ", Y=" & $iMouseY & @CRLF)
                    ConsoleWrite("Face regions count: " & UBound($aFaceRegions) & @CRLF)
                    
                    For $i = 0 To UBound($aFaceRegions) - 1
                        Local $sRegion = $aFaceRegions[$i]
                        ConsoleWrite("Checking Face " & ($i + 1) & ": " & $sRegion & @CRLF)
                        Local $aCoords = StringSplit($sRegion, "|", 2)
                        If IsArray($aCoords) And UBound($aCoords) = 4 Then
                            ConsoleWrite("Face " & ($i + 1) & " region: X1=" & $aCoords[0] & ", Y1=" & $aCoords[1] & ", X2=" & $aCoords[2] & ", Y2=" & $aCoords[3] & @CRLF)
                            If $iMouseX >= $aCoords[0] And $iMouseX <= $aCoords[2] And _
                               $iMouseY >= $aCoords[1] And $iMouseY <= $aCoords[3] Then
                                ConsoleWrite("Face " & ($i + 1) & " SELECTED!" & @CRLF)
                                $iSelectedFace = $i
                                GUICtrlSetData($hSelectionLabel, "Selected: Face " & ($i + 1))
                                ; Update preview to show selection
                                CreateInteractivePreviewImage($sCurrentImagePath, $aDetectedFaces, $iPreviewWidth, $iPreviewHeight, $fScaleX, $fScaleY, $iSelectedFace)
                                ExitLoop
                            EndIf
                        EndIf
                    Next
                EndIf
        EndSwitch
    WEnd
    
    _GDIPlus_ImageDispose($hOriginalImage)
EndFunc

; ------------------------------------------------------------------
;  Create interactive preview image with clickable face regions
; ------------------------------------------------------------------
Func CreateInteractivePreviewImage($sImagePath, $aFaces, $iPreviewWidth, $iPreviewHeight, $fScaleX, $fScaleY, $iSelectedIndex = -1)
    ; Load original image
    Local $hOriginalImage = _GDIPlus_ImageLoadFromFile($sImagePath)
    
    ; Create bitmap for preview
    Local $hBitmap = _GDIPlus_BitmapCreateFromScan0($iPreviewWidth, $iPreviewHeight)
    Local $hGraphics = _GDIPlus_ImageGetGraphicsContext($hBitmap)
    
    ; Draw original image scaled to preview size
    _GDIPlus_GraphicsDrawImageRect($hGraphics, $hOriginalImage, 0, 0, $iPreviewWidth, $iPreviewHeight)
    
    ; Clear face regions array
    Global $aFaceRegions[0]
    
    ; Draw rectangles around detected faces
    For $i = 0 To UBound($aFaces) - 1
        Local $aFace = $aFaces[$i]
        Local $iFaceX = $aFace[0] * $fScaleX
        Local $iFaceY = $aFace[1] * $fScaleY
        Local $iFaceWidth = $aFace[2] * $fScaleX
        Local $iFaceHeight = $aFace[3] * $fScaleY
        
        ; Store face region for click detection (expanded slightly for easier clicking)
        Local $iExpandedX = Max(0, $iFaceX - 5)
        Local $iExpandedY = Max(0, $iFaceY - 5)
        Local $iExpandedWidth = Min($iPreviewWidth - $iExpandedX, $iFaceWidth + 10)
        Local $iExpandedHeight = Min($iPreviewHeight - $iExpandedY, $iFaceHeight + 10)
        
        Local $sRegion = $iExpandedX & "|" & $iExpandedY & "|" & ($iExpandedX + $iExpandedWidth) & "|" & ($iExpandedY + $iExpandedHeight)
        ConsoleWrite("Adding Face " & ($i + 1) & " region: " & $sRegion & @CRLF)
        
        ; Manually add to array since _ArrayAdd might not work properly
        Local $iSize = UBound($aFaceRegions)
        ReDim $aFaceRegions[$iSize + 1]
        $aFaceRegions[$iSize] = $sRegion
        
        ; Choose pen color based on selection
        Local $iPenColor, $iBrushColor
        If $i = $iSelectedIndex Then
            $iPenColor = 0xFF00FF00  ; Green for selected face
            $iBrushColor = 0xFF00FF00
        Else
            $iPenColor = 0xFFFF0000  ; Red for unselected faces
            $iBrushColor = 0xFFFF0000
        EndIf
        
        ; Create pen for face rectangles
        Local $hPen = _GDIPlus_PenCreate($iPenColor, 3)
        
        ; Draw rectangle
        _GDIPlus_GraphicsDrawRect($hGraphics, $iFaceX, $iFaceY, $iFaceWidth, $iFaceHeight, $hPen)
        
        ; Add face number label
        Local $hBrush = _GDIPlus_BrushCreateSolid($iBrushColor)
        Local $hFont = _GDIPlus_FontCreate(_GDIPlus_FontFamilyCreate("Arial"), 14)
        Local $hFormat = _GDIPlus_StringFormatCreate()
        _GDIPlus_GraphicsDrawString($hGraphics, "Face " & ($i + 1), $iFaceX + 5, $iFaceY + 5, $hFont, $hFormat, $hBrush)
        
        _GDIPlus_BrushDispose($hBrush)
        _GDIPlus_FontDispose($hFont)
        _GDIPlus_StringFormatDispose($hFormat)
        _GDIPlus_PenDispose($hPen)
    Next
    
    ; Save preview to temporary file
    Local $sTempFile = @TempDir & "\face_preview_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".jpg"
    _GDIPlus_ImageSaveToFile($hBitmap, $sTempFile)
    
    ; Update the picture control
    GUICtrlSetImage($hPreviewImage, $sTempFile)
    
    ; Cleanup
    _GDIPlus_GraphicsDispose($hGraphics)
    _GDIPlus_BitmapDispose($hBitmap)
    _GDIPlus_ImageDispose($hOriginalImage)
    
    ; Delete temp file on exit
    OnAutoItExitRegister("DeleteTempPreview")
EndFunc

; Helper function to create array
Func _ArrayCreate($v1 = Default, $v2 = Default, $v3 = Default, $v4 = Default, $v5 = Default, $v6 = Default, $v7 = Default, $v8 = Default, $v9 = Default, $v10 = Default)
    Local $aArray[0]
    Return $aArray
EndFunc

; Rest of the functions remain the same as in the previous version...
Func DeleteTempPreview()
    ; Clean up temporary preview files
    Local $aFiles = _FileListToArray(@TempDir, "face_preview_*.jpg", 1)
    If IsArray($aFiles) Then
        For $i = 1 To $aFiles[0]
            FileDelete(@TempDir & "\" & $aFiles[$i])
        Next
    EndIf
EndFunc

; ------------------------------------------------------------------
;  Call Microsoft Face API - FIXED VERSION
; ------------------------------------------------------------------
Func DetectFacesWithMicrosoft($sImagePath)
    ConsoleWrite("Calling Microsoft Face API..." & @CRLF)

    Local $hFile = FileOpen($sImagePath, 16) ; binary
    If $hFile = -1 Then 
        ConsoleWrite("Error: Could not open image file" & @CRLF)
        Return SetError(3, 0, 0)
    EndIf
    Local $bImageData = FileRead($hFile)
    FileClose($hFile)

    ; Build API URL - Remove recognition-related parameters to avoid approval requirements
    Local $sApiUrl = $sEndpoint & "face/v1.0/detect" & _
                    "?returnFaceId=false" & _  ; Disables face recognition
                    "&returnFaceLandmarks=false" & _
                    "&returnFaceAttributes=" & _
                    "&detectionModel=detection_03"  ; Keep newer detection model

    ConsoleWrite("API URL: " & $sApiUrl & @CRLF)

    Local $sResponse = _INetPost($sApiUrl, $bImageData, "application/octet-stream", $sSubscriptionKey)
    If @error Then
        ConsoleWrite("INetPost Error: " & @error & @CRLF)
        Return SetError(2, 0, 0)
    EndIf

    If StringInStr($sResponse, "error") Or $sResponse = "" Then
        ConsoleWrite("API Error Response: " & $sResponse & @CRLF)
        Return SetError(1, 0, 0)
    EndIf

    Local $aFaces = ParseFaceResponse($sResponse)
    Return $aFaces
EndFunc   ;==>DetectFacesWithMicrosoft

; ------------------------------------------------------------------
;  POST via WinHttp (with proper error checking)
; ------------------------------------------------------------------
Func _INetPost($sURL, $vData, $sContentType = "application/octet-stream", $sApiKey = "")
    Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")
    If @error Then 
        ConsoleWrite("Error creating HTTP object: " & @error & @CRLF)
        Return SetError(1, 0, "")
    EndIf
    
    $oHTTP.Open("POST", $sURL, False)
    If @error Then
        ConsoleWrite("Error opening HTTP connection: " & @error & @CRLF)
        Return SetError(2, 0, "")
    EndIf
    
    $oHTTP.SetRequestHeader("Content-Type", $sContentType)
    If @error Then
        ConsoleWrite("Error setting Content-Type header: " & @error & @CRLF)
        Return SetError(3, 0, "")
    EndIf
    
    $oHTTP.SetRequestHeader("Ocp-Apim-Subscription-Key", $sApiKey)
    If @error Then
        ConsoleWrite("Error setting API key header: " & @error & @CRLF)
        Return SetError(4, 0, "")
    EndIf
    
    ; Set longer timeout (30 seconds)
    $oHTTP.SetTimeouts(30000, 30000, 30000, 30000)
    
    $oHTTP.Send($vData)
    If @error Then
        ConsoleWrite("Error sending HTTP request: " & @error & @CRLF)
        Return SetError(5, 0, "")
    EndIf

    ConsoleWrite("HTTP Status: " & $oHTTP.Status & " " & $oHTTP.StatusText & @CRLF)

    If $oHTTP.Status = 200 Then
        Return $oHTTP.ResponseText
    Else
        ConsoleWrite("HTTP Error: " & $oHTTP.Status & " - " & $oHTTP.StatusText & @CRLF)
        ConsoleWrite("Response: " & $oHTTP.ResponseText & @CRLF)
        Return SetError(6, $oHTTP.Status, "")
    EndIf
EndFunc   ;==>_INetPost

; ------------------------------------------------------------------
;  crude JSON extractor
; ------------------------------------------------------------------
Func ParseFaceResponse($sJsonResponse)
    Local $aFaces[0]
    
    ; Validate JSON response
    If StringLeft($sJsonResponse, 1) <> "[" Then
        ConsoleWrite("Invalid JSON response: " & $sJsonResponse & @CRLF)
        Return $aFaces
    EndIf
    
    Local $aFaceBlocks = StringSplit($sJsonResponse, '{"faceRectangle"', 1)
    If $aFaceBlocks[0] < 2 Then 
        ConsoleWrite("No faces found in response" & @CRLF)
        Return $aFaces
    EndIf

    For $i = 2 To $aFaceBlocks[0]
        Local $sFaceBlock = $aFaceBlocks[$i]
        Local $iLeft   = ExtractJsonValue($sFaceBlock, "left")
        Local $iTop    = ExtractJsonValue($sFaceBlock, "top")
        Local $iWidth  = ExtractJsonValue($sFaceBlock, "width")
        Local $iHeight = ExtractJsonValue($sFaceBlock, "height")

        If $iLeft <> "" And $iTop <> "" And $iWidth <> "" And $iHeight <> "" Then
            ReDim $aFaces[UBound($aFaces) + 1]
            Local $aFace[4] = [$iLeft, $iTop, $iWidth, $iHeight]
            $aFaces[UBound($aFaces) - 1] = $aFace

            ConsoleWrite("Detected Face: X=" & $iLeft & ", Y=" & $iTop & _
                        ", W=" & $iWidth & ", H=" & $iHeight & @CRLF)
        EndIf
    Next
    Return $aFaces
EndFunc   ;==>ParseFaceResponse

Func ExtractJsonValue($sJson, $sKey)
    Local $sPattern = '"' & $sKey & '":\s*(\d+)'
    Local $aMatch = StringRegExp($sJson, $sPattern, 1)
    If @error Then Return ""
    Return $aMatch[0]
EndFunc   ;==>ExtractJsonValue

; ------------------------------------------------------------------
;  Crop around the detected face (1:1)
; ------------------------------------------------------------------
Func CropDetectedFace($sImagePath, $sOutputPath, $aFace)
    _GDIPlus_Startup()
    Local $hImage = _GDIPlus_ImageLoadFromFile($sImagePath)
    If @error Then
        _GDIPlus_Shutdown()
        Return False
    EndIf

    Local $iImgWidth  = _GDIPlus_ImageGetWidth($hImage)
    Local $iImgHeight = _GDIPlus_ImageGetHeight($hImage)

    Local $iFaceX      = $aFace[0]
    Local $iFaceY      = $aFace[1]
    Local $iFaceWidth  = $aFace[2]
    Local $iFaceHeight = $aFace[3]

    Local $iCenterX     = $iFaceX + ($iFaceWidth / 2)
    Local $iCenterY     = $iFaceY + ($iFaceHeight / 2)
    Local $iSquareSize  = Max($iFaceWidth, $iFaceHeight) * 1.6 ; 60 % padding

    Local $iCropX = $iCenterX - ($iSquareSize / 2)
    Local $iCropY = $iCenterY - ($iSquareSize / 2)

    $iCropX        = Max(0, $iCropX)
    $iCropY        = Max(0, $iCropY)
    
    ; Fixed: Use Min3 function for three values
    $iSquareSize   = Min3($iSquareSize, $iImgWidth - $iCropX, $iImgHeight - $iCropY)

    If $iSquareSize < 10 Then
        _GDIPlus_Shutdown()
        Return False
    EndIf

    Local $hCropped = _GDIPlus_BitmapCloneArea($hImage, $iCropX, $iCropY, $iSquareSize, $iSquareSize)
    Local $bResult  = _GDIPlus_ImageSaveToFile($hCropped, $sOutputPath)

    _GDIPlus_ImageDispose($hImage)
    _GDIPlus_ImageDispose($hCropped)
    _GDIPlus_Shutdown()

    If $bResult Then ConsoleWrite("Cropped to: " & $iSquareSize & "x" & $iSquareSize & @CRLF)
    Return $bResult
EndFunc   ;==>CropDetectedFace

; ------------------------------------------------------------------
;  Simple center crop
; ------------------------------------------------------------------
Func CenterCrop($sInputPath, $sOutputPath)
    _GDIPlus_Startup()
    Local $hImage = _GDIPlus_ImageLoadFromFile($sInputPath)
    If @error Then 
        _GDIPlus_Shutdown()
        Return False
    EndIf

    Local $iWidth  = _GDIPlus_ImageGetWidth($hImage)
    Local $iHeight = _GDIPlus_ImageGetHeight($hImage)
    
    ; Simple center crop to 1:1 ratio
    Local $iSize = Min($iWidth, $iHeight)
    Local $iX = ($iWidth - $iSize) / 2
    Local $iY = ($iHeight - $iSize) / 2

    $iX = Max(0, $iX)
    $iY = Max(0, $iY)
    
    ; Fixed: Use Min3 function for three values
    $iSize = Min3($iSize, $iWidth - $iX, $iHeight - $iY)

    Local $hCropped = _GDIPlus_BitmapCloneArea($hImage, $iX, $iY, $iSize, $iSize)
    Local $bResult  = _GDIPlus_ImageSaveToFile($hCropped, $sOutputPath)

    _GDIPlus_ImageDispose($hImage)
    _GDIPlus_ImageDispose($hCropped)
    _GDIPlus_Shutdown()
    
    Return $bResult
EndFunc   ;==>CenterCrop

; ------------------------------------------------------------------
;  Min / Max helpers - FIXED VERSIONS
; ------------------------------------------------------------------
Func Max($a, $b)
    Return $a > $b ? $a : $b
EndFunc   ;==>Max

Func Min($a, $b)
    Return $a < $b ? $a : $b
EndFunc   ;==>Min

; NEW: Min function that accepts three parameters
Func Min3($a, $b, $c)
    Local $iMin = $a
    If $b < $iMin Then $iMin = $b
    If $c < $iMin Then $iMin = $c
    Return $iMin
EndFunc   ;==>Min3
