#include <Date.au3>
#include <ColorConstants.au3>
#include <GUIConstantsEx.au3>
#include <Clipboard.au3>
#include <WinAPISys.au3>
#include <WinAPIvkeysConstants.au3>
#Include "win.au3"
Opt("SendKeyDelay", 150)
Opt( "SendKeyDownDelay", 50 )
HotKeySet("{ESC}", "_quit")

Global $sPath_ini = @ScriptDir & "\data.ini"

If FileExists($sPath_ini) Then
   FileDelete($sPath_ini)
EndIf

IniWrite($sPath_ini, "coois", "audiOldOrders","Loading...")
IniWrite($sPath_ini, "coois", "skodaOldOrders", "Loading...")

FileSetAttrib(@ScriptDir & "\data.ini", "+H")

$title0 = "SAP Easy Access"
$title1 = "Material Document List"
$title2 = "Production Order Information System"
$title3 = "Order Info System - Order Headers"
$title4 = "Display Warehouse Stocks of Material on Hand"

$nmb51 = "/nmb51"
$nzmb52 = "/nzmb52"
$ncoois = "/ncoois"
$n = "/n"
$artMin4 = "4000000"
$artMax4 = "4999999"
$cutPlant = "0313"
$sewPlant = "0323"
$movement = "101"
$q3new = "/Q3NEW"
$q3old = "/Q3OLD"
$sk38new = "/SK38NEW"
$sk383old = "/SK38OLD"
$proc = "Z00001"

Global $Q3N1MFB = "/Q3N1MFB"
Global $Q3FB = "/Q3FB"
Global $Q3N1HFB = "/Q3N1HFB" 
Global $Q3N1MFC = "/Q3N1MFC"
Global $Q3FC = "/Q3FC"
Global $Q3RB40 = "/Q3RB40"
Global $Q3RB60 = "/Q3RB60"
Global $Q3RC40 = "/Q3RC40"
Global $Q3RC60 = "/Q3RC60"
Global $SKRB40 = "/SKRB40"
Global $SKRB60 = "/SKRB60"
Global $SKRC = "/SKRC"

Global $audiOldOrders = 'Оновлення даних...'
Global $skodaOldOrders = 'Оновлення даних...'

Global $audiStockOrders = ''
Global $skodaStockOrders = ''

; Todays date
Func _Date($iReturnTime = 1)

   Return @MDAY & '.' & @MON & '.' & @YEAR
EndFunc

; Todays date substract one week
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

Func _select_all()
   _WinAPI_Keybd_Event($VK_CONTROL, 0)
   _WinAPI_Keybd_Event($VK_A, 0)
   _WinAPI_Keybd_Event($VK_A, 2)
   _WinAPI_Keybd_Event($VK_CONTROL, 2)
EndFunc   

Func _clear()
   ControlClick ('','','[ID:1001]','',1)
   _ClipBoard_SetData($n)
   ControlSend('','','[ID:1001]','+{INS}' & '{ENTER}')
EndFunc

Func _check_window($title_exist, $send, $title_need)
   If WinExists($title_exist) Then
      WinActivate($title_exist)
   EndIf
      WinSetState($title_exist,"",@SW_MAXIMIZE)
      _ClipBoard_SetData($send)
      ControlSend($title_exist, "", "[CLASS:Edit; INSTANCE:1; ID:1001]", '+{INS}' & "{ENTER}")
      WinWait( $title_need, "", 0 )
   If WinExists($title_need) Then
      WinActivate($title_need, "" )
   EndIf
   While 1
      WinActivate($title_need)
         If WinActivate($title_need, "" ) Then ExitLoop
   WEnd
EndFunc

_start()

Func _start()
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
   _CrateWinOldOrders()
   Sleep(1000)
   BlockInput($BI_DISABLE)
   _Get_Data_SAP()
   BlockInput($BI_ENABLE)
   _DeleteWinOldOrders()
   _show()
EndFunc

Func _show()
   $hTimer = TimerInit()
   While TimerDiff($hTimer)<25*60*1000
      _CrateWinOldOrders()
      Sleep(30000)
      _DeleteWinOldOrders()
      _CrateWinStockOrders()
      Sleep(30000)
      _DeleteWinStockOrders()
      _ShowDeliveredOrdersAudi()
      Sleep(30000)
      _delDeliveredAudi()
      _ShowDeliveredOrdersSkoda()
      Sleep(30000)
      _delDeliveredSkoda()
      WEnd
   Sleep(500)
   _start()
EndFunc


Func _Coois($point, $layout)
   _check_window($title0, $ncoois, $title2)
   _ClipBoard_SetData($layout)
   Send('{tab}')
   _select_all()
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
	WinWait( $title3, "", 0 )
	While 1
	   WinActivate( $title3, "" )
		If WinExists( $title3 ) Then ExitLoop
	WEnd
	Sleep(3000)
   _ClipBoard_SetData("0")
   Send("{TAB 10}")
   Send('^{END}')
	Sleep(3000)
	Send('^{INS}')
	$point = ClipGet()
	$point=StringRegExpReplace($point,'(\d)\s','\1')
   Sleep(1000)
   _clear()

   Return $point
EndFunc  

Func _Stock($point, $layout)
   _check_window($title0, $nzmb52, $title4)
   _select_all()
   _ClipBoard_SetData($artMin4)
   Send('+{INS}')
   Send('{tab}')
   _ClipBoard_SetData($artMax4)
   Send('+{INS}')
   Send('{tab 2}')
   _select_all()
   _ClipBoard_SetData($sewPlant)
   Send('+{INS}')
   Send("{DOWN}")
   _select_all()
   Send("{DEL}")
   Send("{DOWN}")
   _select_all()
   Send("{DEL}")
	Send("{DOWN 8}")
	Send("+{TAB}")
	Send("{DOWN}")
   Send("{TAB}")
   _select_all()
   _ClipBoard_SetData($layout)
    Send('+{INS}')
   Send("{F8}")
   _ClipBoard_SetData("0")
   While 1
	   WinActivate($title4)
		If ControlGetFocus($title4) == "SAPALVGrid1" Then ExitLoop
   WEnd
   Sleep(1000)
   Send("^{END}")
	Sleep(1000)
	Send('^{INS}')
	$point = ClipGet()
	$point = StringRegExpReplace($point,'(\d)\s','\1')
   $point = StringLeft($point, StringInStr($point, ",") - 1)
   Sleep(1000)
   _clear()
   
   Return $point
EndFunc

Func _GetDataDeliver($point)
   _check_window($title0, $nmb51, $title1)
   _select_all()
   _ClipBoard_SetData($artMin4)
   Send('+{INS}')
   Send("{TAB}")
   _ClipBoard_SetData($artMax4)
   Send('+{INS}')
   Send("{TAB 2}")
   _select_all()
   _ClipBoard_SetData($cutPlant)
   Send('+{INS}')
   Send("{DOWN}")
   _select_all()
   _ClipBoard_SetData($cutPlant)
   Send('+{INS}')
   Send("{DOWN 4}")
   _select_all()
   _ClipBoard_SetData($movement)
   Send('+{INS}')
   Send("{DOWN 2}")
   _select_all()
   Send("{DEL}")
   Send("{DOWN 3}")
   _select_all()
   _ClipBoard_SetData(_Date())
   Send('+{INS}')
   Send("{DOWN 4}")
   _select_all()
   _ClipBoard_SetData($point)
   Send('+{INS}')
   Send("{F8}")
   _ClipBoard_SetData("0")
   While 1
      WinActivate($title1)
      If ControlGetFocus($title1) == "SAPALVGrid1" Then ExitLoop
   WEnd
   Sleep(1000)
   Send("^{END}")
	Sleep(1000)
	Send('^{INS}')
	$point = ClipGet()
   $point = StringRegExpReplace($point,'(\d)\s','\1')
   If $point = '' Then
      $point = "0"
   EndIf
   Sleep(1000)
   _clear()
   
   Return $point
EndFunc

Func _Get_Data_SAP()
   IniWrite($sPath_ini, "coois", "audiOldOrders", _Coois($audiOldOrders, $q3old))
   WinWaitActive($title0)
   IniWrite($sPath_ini, "coois", "skodaOldOrders", _Coois($skodaOldOrders, $sk383old))
   WinWaitActive($title0)
   IniWrite($sPath_ini, "Stock", "audiStockOrders", _Stock($audiStockOrders, $q3new))
   WinWaitActive($title0)
   IniWrite($sPath_ini, "Stock", "skodaStockOrders", _Stock($skodaStockOrders, $sk38new))
   WinWaitActive($title0)
   IniWrite($sPath_ini, "Q3", "Q3N1MFB", _GetDataDeliver($Q3N1MFB))
   WinWaitActive($title0)
   IniWrite($sPath_ini, "Q3", "Q3FB", _GetDataDeliver($Q3FB))
   WinWaitActive($title0)
   IniWrite($sPath_ini, "Q3", "Q3N1HFB", _GetDataDeliver($Q3N1HFB))
   WinWaitActive($title0)
   IniWrite($sPath_ini, "Q3", "Q3N1MFC", _GetDataDeliver($Q3N1MFC))
   WinWaitActive($title0)
   IniWrite($sPath_ini, "Q3", "Q3FC", _GetDataDeliver($Q3FC))
   WinWaitActive($title0)
   IniWrite($sPath_ini, "Q3", "Q3RB40", _GetDataDeliver($Q3RB40))
   WinWaitActive($title0)
   IniWrite($sPath_ini, "Q3", "Q3RB60", _GetDataDeliver($Q3RB60))
   WinWaitActive($title0)
   IniWrite($sPath_ini, "Q3", "Q3RC40", _GetDataDeliver($Q3RC40))
   WinWaitActive($title0)
   IniWrite($sPath_ini, "Q3", "Q3RC60", _GetDataDeliver($Q3RC60))
   WinWaitActive($title0)
   IniWrite($sPath_ini, "SK38", "SKRB40", _GetDataDeliver($SKRB40))
   WinWaitActive($title0)
   IniWrite($sPath_ini, "SK38", "SKRB60", _GetDataDeliver($SKRB60))
   WinWaitActive($title0)
   IniWrite($sPath_ini, "SK38", "SKRC", _GetDataDeliver($SKRC))
   WinWaitActive($title0)
EndFunc

Func _quit()
   FileDelete($sPath_ini)
	Exit
EndFunc