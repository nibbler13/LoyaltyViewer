#include <Array.au3>
#include <FileConstants.au3>
#include "XML.au3"
#include <GuiListView.au3>
#include <ButtonConstants.au3>
#include <DateTimeConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <Date.au3>
#include <FontConstants.au3>
#include <ColorConstants.au3>
#include <IE.au3>

Local $nScreenWidth = 1024;@DesktopWidth
Local $nScreenHeight = 768;@DesktopHeight
Local $sResourcePath = @ScriptDir & "\" & "Resources\"
Local $sFileName = $sResourcePath & "PicMainBackground_*.jpg"

Local $nFontSize = $nScreenHeight / 24
Local $nFontWeight = 200
Local $sFontName = "Franklin Gothic"
Local $sFontNameSub = "Franklin Gothic Book"

Local $nImageWidth = 1180
Local $nImageHeight = 671
Local $nImageRatio = $nImageHeight / $nImageWidth

$nImageWidth = $nScreenWidth * ($nImageWidth / 1920)
$nImageHeight = $nImageWidth * $nImageRatio

Local $sTextMain = "Сегодня нас рекомендуют более *% пациентов!"
Local $sTextSub = "ВНИМАНИЕ! Работают эвакуаторы, пользуйтесь подземной парковкой."

Global Const $HTTP_STATUS_OK = 200


$oMyError = ObjEvent("AutoIt.Error","MyErrFunc")





Local $hGui = GUICreate("SelfChecking", $nScreenWidth, $nScreenHeight, 0, 0, $WS_POPUP);, $WS_EX_TOPMOST)
GUISetBkColor($COLOR_WHITE)

GUISetState(@SW_SHOW)

Local $oIEMap = _IECreateEmbedded()
GUICtrlCreateObj($oIEMap, 0, 0, $nScreenWidth * 0.3, $nScreenHeight * 0.3)
_IENavigate($oIEMap, "https://yandex.ru/maps/213/moscow/?lang=ru&ncrnd=2813&l=trf%2Ctrfe&ll=37.571648%2C55.762964&z=11")

Local $oIE = _IECreateEmbedded()
GUICtrlCreateObj($oIE, $nScreenWidth * 0.7, 0, $nScreenWidth * 0.3, $nScreenHeight * 0.3)
_IENavigate($oIE, "https://www.gismeteo.ru/ajax/print/4368/weekly/")



Local $hImage = GUICtrlCreatePic(StringReplace($sFileName, "*", "0"), $nScreenWidth / 2 - $nImageWidth / 2, _
	$nScreenHeight / 2 - $nImageHeight / 2, $nImageWidth, $nImageHeight)
Local $hLabelMain = GUICtrlCreateLabel("", 0, $nScreenHeight - $nFontSize * 4, $nScreenWidth, $nFontSize * 1.5, $SS_CENTER)
GUICtrlSetBkColor(-1, 0xedf3e9)
GUICtrlSetColor(-1, 0xb50708)
GUICtrlSetFont(-1, $nFontSize, $nFontWeight, 0, $sFontNameSub, $CLEARTYPE_QUALITY)

Local $hLabelSub = GUICtrlCreateLabel($sTextSub, 0, $nScreenHeight - $nFontSize * 2, $nScreenWidth, $nFontSize * 1.5, $SS_CENTER)
GUICtrlSetColor(-1, 0x5a4e50)
GUICtrlSetFont(-1, $nFontSize * 0.9, $nFontWeight, 0, $sFontNameSub, $CLEARTYPE_QUALITY)

Local $prevData = 0




While 1
	If GUIGetMsg() = $GUI_EVENT_CLOSE Then Exit

	Sleep(500)

	ReloadData()

	Sleep(20 * 1000)

	_IEAction($oIE, "refresh")
	_IEAction($oIEMap, "refresh")
WEnd



Func ReloadData()
	Local $aTmpDate, $aTmpTime
	_DateTimeSplit(_NowCalc(), $aTmpDate, $aTmpTime)
	Local $strDateNow = $aTmpDate[3] & "." & $aTmpDate[2] & "." & $aTmpDate[1]

	Local $strUrl = ""
	Local $strReportID = "EMPLOYEE"
	Local $strDtBegin = $strDateNow & " 00:00:00"
	Local $strDtEnd = $strDateNow & " 23:59:59"
	Local $strUseUTC = "0"
	Local $strLoginName = ""
	Local $strPassword = ""
	Local $nQuestionID = 152
	Local $strProxy = ""

	Local $strBody = "LoginName=" & $strLoginName & _
					 "&Password=" & $strPassword & _
					 "&QuestionID=" & $nQuestionID & _
					 "&ReportID=" & $strReportID & _
					 "&Begin=" & $strDtBegin & _
					 "&End=" & $strDtEnd & _
					 "&UseUTC=" & $strUseUTC

;~ 	ConsoleWrite($strBody & @CRLF)

;~ 	MsgBox(0, "", $strBody)

	Local $strXmlResponse = HttpPost($strUrl, $strBody, $strProxy)

	Local $strFileName = @ScriptDir & "\response.xml"
	Local $hFile = FileOpen($strFileName, BitOR($FO_OVERWRITE, $FO_ANSI))
	FileWrite($hFile, $strXmlResponse)
	FileClose($hFile)

	Local $resultArray = ParseXmlFileToArray($strFileName)
;~ 	_ArrayDisplay($resultArray)

	Local $result[0][12]
	If IsArray($resultArray) Then
;~ 		ConsoleWrite("Array" & @CRLF)
		For $i = 1 To UBound($resultArray, $UBOUND_ROWS) - 1
			Local $tmpArray[1][12]
			For $x = 0 To 11
				$tmpArray[0][$x] = $resultArray[$i + $x][3]
			Next
			_ArrayAdd($result, $tmpArray)
			$i += 11
		Next
	Else
;~ 		ConsoleWrite("not Array" & @CRLF)
		Local $tmp[1][12]
		$tmp[0][0] = "Нет данных"
		_ArrayAdd($result, $tmp)
	EndIf

;~ 	_ArrayDisplay($result)
	Local $nRows = Ubound($result, $UBOUND_ROWS)
	If Not $nRows Then Return

	Local $nMark = Number($result[$nRows - 1][7], $NUMBER_DOUBLE) + Number($result[$nRows - 1][8], $NUMBER_DOUBLE)
	Local $sReplace = ""

	If $nMark < 75 Then $sReplace = 75
	For $i = 75 To 95 Step 5
		If $nMark >= $i Then $sReplace = $i
	Next

	If $prevData = $sReplace Then Return

	$prevData = $sReplace

;~ 	ConsoleWrite($nMark & " " & $sReplace & @CRLF)

	GUICtrlSetData($hLabelMain, $sReplace ? StringReplace($sTextMain, "*", $sReplace) : "")
	GUICtrlSetImage($hImage, StringReplace($sFileName, "*", $sReplace))
EndFunc




Func ParseXmlFileToArray($strFileName)
	Local $oXMLDoc = _XML_CreateDOMDocument(Default)
	If @error Then Return

	Local $oXMLDOM_EventsHandler = ObjEvent($oXMLDoc, "XML_DOM_EVENT_")

	_XML_Load($oXMLDoc, $strFileName)
	If @error Then Return

	Local $sXmlAfterTidy = _XML_TIDY($oXMLDoc)
	If @error Then Return

	Local $oNodesColl = _XML_SelectNodes($oXMLDoc, "//Rows/Row/Cell")
	If @error Then Return

	Local $aNodesColl = _XML_Array_GetNodesProperties($oNodesColl)
	If @error Then Return

	Return($aNodesColl)
EndFunc    ;==>Example_1__XML_SelectNodes



Func HttpPost($sURL, $sData = "", $strProxy = "")
	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")
	If $strProxy Then $oHTTP.SetProxy(2, $strProxy)

	$oHTTP.Open("POST", $sURL, False)
	If (@error) Then Return SetError(1, 0, 0)

	$oHTTP.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
	$oHTTP.SetRequestHeader("RequestType", "GetXmlLoyaltyReport")
	$oHTTP.SetRequestHeader("Content-Length", StringLen($sData))

	$oHTTP.Send($sData)
	If (@error) Then Return SetError(2, 0, 0)
	If ($oHTTP.Status <> $HTTP_STATUS_OK) Then Return SetError(3, 0, 0)

	Return SetError(0, 0, $oHTTP.ResponseText)
EndFunc


Func MyErrFunc()
  Msgbox(0,"AutoItCOM Test","We intercepted a COM Error !"    & @CRLF  & @CRLF & _
             "err.description is: " & @TAB & $oMyError.description  & @CRLF & _
             "err.windescription:"   & @TAB & $oMyError.windescription & @CRLF & _
             "err.number is: "       & @TAB & hex($oMyError.number,8)  & @CRLF & _
             "err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @CRLF & _
             "err.scriptline is: "   & @TAB & $oMyError.scriptline   & @CRLF & _
             "err.source is: "       & @TAB & $oMyError.source       & @CRLF & _
             "err.helpfile is: "       & @TAB & $oMyError.helpfile     & @CRLF & _
             "err.helpcontext is: " & @TAB & $oMyError.helpcontext _
            )
Endfunc