#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=C:\Users\evhen\OneDrive\Рабочий стол\Logo_Bader\BADER.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <AutoItConstants.au3>
#include <TrayConstants.au3>
#include <MemoryConstants.au3>
#include <MsgBoxConstants.au3>
#include <Date.au3>
#include <Clipboard.au3>
#include <StaticConstants.au3>
#include <FontConstants.au3>
#include <GUIConstants.au3>
#include <EditConstants.au3>
#include <WinAPIEx.au3>
#include <ColorConstants.au3>
#include <WinAPI.au3>
#Include <GuiButton.au3>
#include <WinAPISys.au3>
#include <WinAPIvkeysConstants.au3>
#Include "window.au3"

AutoItSetOption ( "SendKeyDelay", 200 )
AutoItSetOption ( "SendKeyDownDelay", 200 )
AutoItSetOption ( "WinSearchChildren", 1 )
AutoItSetOption ( "WinTitleMatchMode", 4 )
AutoItSetOption ( "WinWaitDelay", 500 )

HotKeySet("{ESC}", "_quit")

$sFont = "Arial"
$title0 = "SAP Easy Access"
$title1 = "Material Document List"
$title2 = "Production Order Information System"
$title3 = "Order Info System - Order Headers"
$title4 = "Display Warehouse Stocks of Material on Hand"

$vData = "0"
$timeToSleep = "30000"

$nmb51 = "/nmb51"
$artMin4 = "4000000"
$artMax4 = "4999999"
$cutPlant = "0313"
$movement = "101"
$nzmb52 = "/nzmb52"
$ncoois = "/ncoois"
$n = "/n"
$sewPlant = "0323"
$q3new = "/q3new"
$q3old = "/q3old"
$sk38new = "/sk38new"
$sk383old = "/sk38old"
$proc = "z00001"


Global $Q3N1MFB = "0"
Global $Q3FB = "0"
Global $Q3N1HFB = "0" 
Global $Q3N1MFC = "0"
Global $Q3FC = "0"
Global $Q3RB40 = "0"
Global $Q3RB60 = "0"
Global $Q3RC40 = "0"
Global $Q3RC60 = "0"
Global $SKRB40 = "0"
Global $SKRB60 = "0"
Global $SKRC = "0"

Global $audiOldOrders = 'Оновлення даних...'
Global $skodaOldOrders = 'Оновлення даних...'
Global $audiNewOrders = 'Оновлення даних...'
Global $skodaNewOrders = 'Оновлення даних...'
Global $audiStockOrders = 'Оновлення даних...'
Global $skodaStockOrders = 'Оновлення даних...'

;~  Begin run

Func _Date($iReturnTime = 1)
    Return @MDAY & '.' & @MON & '.' & @YEAR
EndFunc

Dim $aDate = StringSplit(@MDAY & '/' & @MON & '/' & @YEAR, '/')

$sDay = _DateDayOfWeekEx(_DateToDayOfWeek($aDate[3], $aDate[2], $aDate[1]), 0)

Func _DateDayOfWeekEx($iDayNum, $iShort = 0)
	Local Const $aDayOfWeek[8] = ["", "0", "1", "2", "3", "4", "5", "6"]
	Local Const $aDayOfWeek_Rus[8] = ["", "0", "1", "2", "3", "4", "5", "6"]

	Select
		Case Not StringIsInt($iDayNum) Or Not StringIsInt($iShort) Or $iDayNum < 1 Or $iDayNum > 7
			Return SetError(1, 0, "")
		Case Else
			Local $sRet = $aDayOfWeek[$iDayNum]
			If @OSLang = 0419 Then $sRet = $aDayOfWeek_Rus[$iDayNum]

			If $iShort Then
				$sRet = StringLeft($sRet, 3)
			EndIf

			Return $sRet
	EndSelect
EndFunc

$sDatePast = SubstractDate($sDay)

Func SubstractDate($i)
    If $i == 1 Then
       $sDate = _DateAdd('d', -7, _NowCalcDate())
       $aArray = StringSplit($sDate, "/", 2)
       $sDatePast = $aArray[2] &"."& $aArray[1] &"."& StringRight($aArray[0], 4)
    ElseIf $i == 2 Then
       $sDate = _DateAdd('d', -7, _NowCalcDate())
       $aArray = StringSplit($sDate, "/", 2)
       $sDatePast = $aArray[2] &"."& $aArray[1] &"."& StringRight($aArray[0], 4)
    ElseIf $i == 3 Then
       $sDate = _DateAdd('d', -7, _NowCalcDate())
       $aArray = StringSplit($sDate, "/", 2)
       $sDatePast = $aArray[2] &"."& $aArray[1] &"."& StringRight($aArray[0], 4)
    ElseIf $i == 4 Then
       $sDate = _DateAdd('d', -7, _NowCalcDate())
       $aArray = StringSplit($sDate, "/", 2)
       $sDatePast = $aArray[2] &"."& $aArray[1] &"."& StringRight($aArray[0], 4)
    ElseIf $i == 5 Then
       $sDate = _DateAdd('d', -7, _NowCalcDate())
       $aArray = StringSplit($sDate, "/", 2)
       $sDatePast = $aArray[2] &"."& $aArray[1] &"."& StringRight($aArray[0], 4)
    ElseIf $i == 6 Then
       $sDate = _DateAdd('d', -5, _NowCalcDate())
       $aArray = StringSplit($sDate, "/", 2)
       $sDatePast = $aArray[2] &"."& $aArray[1] &"."& StringRight($aArray[0], 4)
    ElseIf $i == 0 Then
       $sDate = _DateAdd('d', -6, _NowCalcDate())
       $aArray = StringSplit($sDate, "/", 2)
       $sDatePast = $aArray[2] &"."& $aArray[1] &"."& StringRight($aArray[0], 4)
 EndIf

 Return $sDatePast
EndFunc


Func _SetColor($data, $min, $max, $labelID)
	If $data > $max Then
        GUICtrlSetBkColor($labelID, 0x0011DD15)
     ElseIf $data >= $min And $data <= $max Then
        GUICtrlSetBkColor($labelID, $COLOR_YELLOW)
     ElseIf $data < $min Then
        GUICtrlSetBkColor($labelID, $COLOR_RED)
     EndIf
EndFunc



_start()

Func _start()
    _ShowOldOrders()
    While 1
	    WinActivate($title0)
	        If WinActive($title0) Then ExitLoop
        WEnd
        $Q3N1MFB = "/Q3N1MFB"
        $Q3FB = "/Q3FB"
        $Q3N1HFB = "/Q3N1HFB" 
        $Q3N1MFC = "/Q3N1MFC"
        $Q3FC = "/Q3FC"
        $Q3RB40 = "/Q3RB40"
        $Q3RB60 = "/Q3RB60"
        $Q3RC40 = "/Q3RC40"
        $Q3RC60 = "/Q3RC60"
        $SKRB40 = "/SKRB40"
        $SKRB60 = "/SKRB60"
        $SKRC = "/SKRC"
        Sleep(1000)
        _chooseData($sDatePast)
        _delOld()
        _show()
EndFunc

Func _show()
    $hTimer = TimerInit()
    While TimerDiff($hTimer)<15*60*1000
       _ShowOldOrders()
       Sleep($timeToSleep)
       _delOld()
       _ShowStockOrders()
       Sleep($timeToSleep)
       _delStock()
       _ShowDeliveredOrdersAudi()
       Sleep($timeToSleep)
       _delDeliveredAudi()
       _ShowDeliveredOrdersSkoda()
       Sleep($timeToSleep)
       _delDeliveredSkoda()
    WEnd
    Sleep(500)
    _start()
 EndFunc

 Func _clear()
    ControlClick ('','','[ID:1001]','',1)
    _ClipBoard_SetData($n)
    ControlSend('','','[ID:1001]','+{INS}' & '{ENTER}')
EndFunc

Func _GetDataDeliver($point)
    _WinAPI_Keybd_Event($VK_CONTROL, 0)
    _WinAPI_Keybd_Event($VK_A, 0)
    _WinAPI_Keybd_Event($VK_A, 2)
    _WinAPI_Keybd_Event($VK_CONTROL, 2)
    _ClipBoard_SetData($artMin4)
    Send('+{INS}')
    Send("{TAB}")
    _ClipBoard_SetData($artMax4)
    Send('+{INS}')
    Send("{TAB 2}")
    _WinAPI_Keybd_Event($VK_CONTROL, 0)
    _WinAPI_Keybd_Event($VK_A, 0)
    _WinAPI_Keybd_Event($VK_A, 2)
    _WinAPI_Keybd_Event($VK_CONTROL, 2)
    _ClipBoard_SetData($cutPlant)
    Send('+{INS}')
    Send("{DOWN}")
    _WinAPI_Keybd_Event($VK_CONTROL, 0)
    _WinAPI_Keybd_Event($VK_A, 0)
    _WinAPI_Keybd_Event($VK_A, 2)
    _WinAPI_Keybd_Event($VK_CONTROL, 2)

    _ClipBoard_SetData($cutPlant)
    Send('+{INS}')
    Send("{DOWN 4}")
    _WinAPI_Keybd_Event($VK_CONTROL, 0)
    _WinAPI_Keybd_Event($VK_A, 0)
    _WinAPI_Keybd_Event($VK_A, 2)
    _WinAPI_Keybd_Event($VK_CONTROL, 2)

    _ClipBoard_SetData($movement)
    Send('+{INS}')
    Send("{DOWN 2}")
    _WinAPI_Keybd_Event($VK_CONTROL, 0)
    _WinAPI_Keybd_Event($VK_A, 0)
    _WinAPI_Keybd_Event($VK_A, 2)
    _WinAPI_Keybd_Event($VK_CONTROL, 2)

    Send("{DEL}")
    Send("{DOWN 3}")
    _WinAPI_Keybd_Event($VK_CONTROL, 0)
    _WinAPI_Keybd_Event($VK_A, 0)
    _WinAPI_Keybd_Event($VK_A, 2)
    _WinAPI_Keybd_Event($VK_CONTROL, 2)

    _ClipBoard_SetData(_Date())
    Send('+{INS}')
    Send("{DOWN 4}")
    _WinAPI_Keybd_Event($VK_CONTROL, 0)
    _WinAPI_Keybd_Event($VK_A, 0)
    _WinAPI_Keybd_Event($VK_A, 2)
    _WinAPI_Keybd_Event($VK_CONTROL, 2)

    _ClipBoard_SetData($point)
    Send('+{INS}')
    Send("{F8}")
	While 1
	    WinActivate($title1)
		  If ControlGetFocus($title1) == "SAPALVGrid1" Then ExitLoop
    WEnd
	Send("^{END}")
	Sleep(1000)
    Send('^{c}')
    $point = ClipGet()
    $point=StringRegExpReplace($point,'(\d)\s','\1')
    if $point = "" Then 
        $point = "0"
    EndIf
    Sleep(1000)
    _ClipBoard_SetData($nmb51)
    ControlSend($title1, "", "[CLASS:Edit; INSTANCE:1; ID:1001]", '+{INS}' & "{ENTER}")
	While 1
	   WinActivate($title1)
	      If ControlGetFocus($title1) <> "SAPALVGrid1" Then ExitLoop
    WEnd
    Sleep(1000)
    Return $point
EndFunc

Func _GetDataStock($point, $layout)
    WinWait( $title0, "", 0 )
	If WinExists( $title0 ) Then
		WinActivate( $title0 )
    EndIf
       WinSetState($title0,"",@SW_MAXIMIZE)
       _ClipBoard_SetData($nzmb52)
	   ControlSend($title0, "", "[CLASS:Edit; INSTANCE:1; ID:1001]", '+{INS}' & "{ENTER}")
	   WinWait( $title4, "", 0 )
	If WinExists( $title4 ) Then
		WinActivate( $title4, "" )
	EndIf
	While 1
	   WinActivate($title4)
		   If WinActivate( $title4, "" ) Then ExitLoop
	WEnd
	Send("{BACKSPACE}")
    _WinAPI_Keybd_Event($VK_CONTROL, 0)
    _WinAPI_Keybd_Event($VK_A, 0)
    _WinAPI_Keybd_Event($VK_A, 2)
    _WinAPI_Keybd_Event($VK_CONTROL, 2)
    _ClipBoard_SetData($artMin4)
    Send('+{INS}')
	Send("{TAB}")
    _ClipBoard_SetData($artMax4)
    Send('+{INS}')
	Send("{TAB 2}")
    _ClipBoard_SetData($sewPlant)
    Send('+{INS}')
	Send("{DOWN}")
    _WinAPI_Keybd_Event($VK_CONTROL, 0)
    _WinAPI_Keybd_Event($VK_A, 0)
    _WinAPI_Keybd_Event($VK_A, 2)
    _WinAPI_Keybd_Event($VK_CONTROL, 2)
	Send("{DEL}")
	Send("{DOWN 9}")
	Send("+{TAB}")
	Send("{DOWN}")
	Send("{TAB}")
    _WinAPI_Keybd_Event($VK_CONTROL, 0)
    _WinAPI_Keybd_Event($VK_A, 0)
    _WinAPI_Keybd_Event($VK_A, 2)
    _WinAPI_Keybd_Event($VK_CONTROL, 2)
    _ClipBoard_SetData($layout)
    Send('+{INS}')
	Send("{F8}")
	Sleep(5000)
	While 1
	   WinActivate($title4)
		  If ControlGetFocus($title4) == "SAPALVGrid1" Then ExitLoop
	WEnd
    Send("^{END}")
    _ClipBoard_SetData($vData)
	Sleep(3000)
	Send('^{c}')
	$point = ClipGet()
	$point = StringRegExpReplace($point,'(\d)\s','\1')
    $point = StringLeft($point, StringInStr($point, ",") - 1)
    _clear()
    Sleep(1000)
    Return $point
EndFunc

Func _GetDataCooisOldOrders($point, $layout)
    WinWait( $title0, "", 0 )
	If WinExists( $title0 ) Then
		WinActivate( $title0 )
    EndIf
       WinSetState($title0,"",@SW_MAXIMIZE)
       _ClipBoard_SetData($ncoois)
	   ControlSend($title0, "", "[CLASS:Edit; INSTANCE:1; ID:1001]", '+{INS}' & "{ENTER}")
	   WinWait( $title2, "", 0 )
	If WinExists( $title2 ) Then
		WinActivate( $title2, "" )
	EndIf
	While 1
	   WinActivate($title2)
		   If WinActivate( $title2, "" ) Then ExitLoop
	WEnd
	Sleep(1000)
	Send('{tab}')
    _WinAPI_Keybd_Event($VK_CONTROL, 0)
    _WinAPI_Keybd_Event($VK_A, 0)
    _WinAPI_Keybd_Event($VK_A, 2)
    _WinAPI_Keybd_Event($VK_CONTROL, 2)
    _ClipBoard_SetData($layout)
    Send('+{INS}')
    Send('{tab 6}')
    _ClipBoard_SetData($artMin4)
    Send('+{INS}')
    Send('{tab}')
    _ClipBoard_SetData($artMax4)
    Send('+{INS}')
    Send('{tab 2}')
    _ClipBoard_SetData($cutPlant)
    Send('+{INS}')
	Send("{DOWN 11}")
    _ClipBoard_SetData($proc)
    Send('+{INS}')
	Send("{DOWN 18}")
    Send("{TAB 2}")
    _ClipBoard_SetData($sDatePast)
    Send('+{INS}')
	Send('{F8}')
	Sleep(5000)
	WinWait( $title3, "", 0 )
	While 1
	   WinActivate( $title3, "" )
		  If WinExists( $title3 ) Then ExitLoop
	WEnd
	Sleep(3000)
	Send("{TAB 10}")
    Send('^{END}')
    _ClipBoard_SetData($vData)
	Sleep(3000)
	Send('^{c}')
	$point = ClipGet()
	$point=StringRegExpReplace($point,'(\d)\s','\1')
    Sleep(3000)
    _clear()
    Return $point
EndFunc

Func _chooseData($sDatePast)
	WinWait( $title0, "", 0 )
	If WinExists( $title0 ) Then
		WinActivate( $title0 )
    EndIf
        _ClipBoard_SetData($nmb51)
	    WinSetState($title0,"",@SW_MAXIMIZE)
	    ControlSend($title0, "", "[CLASS:Edit; INSTANCE:1; ID:1001]",  '+{INS}' & "{ENTER}")
	    WinWait( $title1, "", 0 )
	If WinExists( $title1 ) Then
	    WinActivate( $title1, "" )
	EndIf
	While 1
	   WinActivate($title1)
		   If WinActivate( $title1, "" ) Then ExitLoop
    WEnd
    $Q3N1MFB = _GetDataDeliver($Q3N1MFB)
    $Q3FB = _GetDataDeliver($Q3FB)
    $Q3N1HFB = _GetDataDeliver($Q3N1HFB)
    $Q3N1MFC = _GetDataDeliver($Q3N1MFC)
    $Q3FC = _GetDataDeliver($Q3FC)
    $Q3RB40 = _GetDataDeliver($Q3RB40)
    $Q3RB60 = _GetDataDeliver($Q3RB60)
    $Q3RC40 = _GetDataDeliver($Q3RC40)
    $Q3RC60 = _GetDataDeliver($Q3RC60)
    $SKRB40 = _GetDataDeliver($SKRB40)
    $SKRB60 = _GetDataDeliver($SKRB60)
    $SKRC = _GetDataDeliver($SKRC)
    _clear()
    $audiStockOrders = _GetDataStock($audiStockOrders, $q3new)
    $skodaStockOrders = _GetDataStock($skodaStockOrders, $sk38new)
    $audiOldOrders = _GetDataCooisOldOrders($audiOldOrders, $q3old)
    $skodaOldOrders = _GetDataCooisOldOrders($skodaOldOrders, $sk383old)
EndFunc

Func _quit()
	Exit
EndFunc