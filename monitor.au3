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

AutoItSetOption ( "SendKeyDelay", 100 )
AutoItSetOption ( "SendKeyDownDelay", 100 )
AutoItSetOption ( "WinSearchChildren", 1 )
AutoItSetOption ( "WinTitleMatchMode", 4 )
AutoItSetOption ( "WinWaitDelay", 500 )

HotKeySet("^{q}", "_quit")

Func _quit()
	Exit
 EndFunc

Func _Date($iReturnTime = 1)
    Return @MDAY & '.' & @MON & '.' & @YEAR
EndFunc

Global $title0 = "SAP Easy Access"
Global $title1 = "Material Document List"
Global $title2 = "Production Order Information System"
Global $title3 = "Order Info System - Order Headers"
Global $title4 = "Display Warehouse Stocks of Material on Hand"

$nzmb52 = "/nzmb52{enter}"
$nmb51 = "/nmb51{enter}"
$ncoois = "/ncoois{enter}"
$n = "/n{enter}"
$art4min = "4000000"
$art4max = "4999999"
$plant = "0313"
$plant1 = "0323"
$movement = "101"
$q3new = "/q3new"
$q3old = "/q3old"
$sk38new = "/sk38new"
$sk383old = "/sk38old"
$proc = "z00001"

Global $audiOldOrders = 'Оновлення даних...'
Global $skodaOldOrders = 'Оновлення даних...'
Global $audiNewOrders = 'Оновлення даних...'
Global $skodaNewOrders = 'Оновлення даних...'
Global $audiStockOrders = 'Оновлення даних...'
Global $skodaStockOrders = 'Оновлення даних...'

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

_start()

Func _show()
   $hTimer = TimerInit()
   While TimerDiff($hTimer)<5*60*1000
	  _ShowOldOrders()
	  Sleep(30000)
	  _delOld()
	  _ShowStockOrders()
	  Sleep(30000)
	  _delStock()
   WEnd
	  Sleep(500)
	  _start()
EndFunc

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

Func _start()
   _ShowOldOrders()
   While 1
	   WinActivate($title0)
	   If WinActive($title0) Then ExitLoop
		WEnd
   _chooseData($sDatePast)
   _delOld()
   _show()
EndFunc

Func _delOld()
   GUIDelete($mainWindow)
EndFunc

Func _delStock()
   GUIDelete($mainWindowStock)
EndFunc

Func _chooseData($sDatePast)
	;~ START GET DATA
	WinWait( $title0, "", 0 )
	If WinExists( $title0 ) Then
		WinActivate( $title0 )
	EndIf
	   WinSetState($title0,"",@SW_MAXIMIZE)
	   ControlSend($title0, "", "[CLASS:Edit; INSTANCE:1; ID:1001]", $nzmb52 & "{ENTER}")
	   WinWait( $title4, "", 0 )
	If WinExists( $title4 ) Then
		WinActivate( $title4, "" )
	EndIf
	While 1
	   WinActivate($title4)
		   If WinActivate( $title4, "" ) Then ExitLoop
	WEnd
	$vData = "0"
	_ClipBoard_SetData($vData)
	Send("{BACKSPACE}")
	Send("^{a}")
	Send($art4min)
	Send("{TAB}")
	Send($art4max)
	Send("{TAB 2}")
	Send($plant1)
	Send("{DOWN}")
	Send("^{a}")
	Send("{DEL}")
	Send("{DOWN 9}")
	Send("+{TAB}")
	Send("{DOWN}")
	Send("{TAB}")
	Send("^{a}")
	Send($q3new)
	Send("{F8}")
	Sleep(5000)
	While 1
	   WinActivate($title4)
		  If ControlGetFocus($title4) == "SAPALVGrid1" Then ExitLoop
	WEnd
	Send("^{END}")
	Sleep(3000)
	Send('^{c}')
	$audiStockOrders = ClipGet()
	$audiStockOrders = StringRegExpReplace($audiStockOrders,'(\d)\s','\1')
	$audiStockOrders = StringLeft($audiStockOrders, StringInStr($audiStockOrders, ",") - 1)
	Sleep(3000)
	;~ SKODA - stock level
	ControlSend($title4, "", "[CLASS:Edit; INSTANCE:1; ID:1001]", $nzmb52 & "{ENTER}")
	While 1
	   WinActivate($title4)
	      If ControlGetFocus($title4) <> "SAPALVGrid1" Then ExitLoop
	WEnd
	_ClipBoard_SetData($vData)
	Send("{BACKSPACE}")
	Send("^{a}")
	Send($art4min)
	Send("{TAB}")
	Send($art4max)
	Send("{TAB 2}")
	Send($plant1)
	Send("{DOWN}")
	Send("^{a}")
	Send("{DEL}")
	Send("{DOWN 9}")
	Send("+{TAB}")
	Send("{DOWN}")
	Send("{TAB}")
	Send("^{a}")
	Send($sk38new)
	Send("{F8}")
	Sleep(5000)
	While 1
	   WinActivate($title4)
		  If ControlGetFocus($title4) == "SAPALVGrid1" Then ExitLoop
	WEnd
	Send("^{END}")
	Sleep(3000)
	Send('^{c}')
	$skodaStockOrders = ClipGet()
	$skodaStockOrders = StringRegExpReplace($skodaStockOrders,'(\d)\s','\1')
	$skodaStockOrders = StringLeft($skodaStockOrders, StringInStr($skodaStockOrders, ",") - 1)
	Sleep(3000)
	;~ COOIS audi old oeders
	ControlSend($title4, "", "[CLASS:Edit; INSTANCE:1; ID:1001]", $ncoois & "{ENTER}")
	WinWait( $title2, "", 0 )
	While 1
	   WinActivate( $title2, "" )
		  If WinExists( $title2 ) Then ExitLoop
	WEnd
	_ClipBoard_SetData($vData)
	Sleep(1000)
	Send('{tab}')
	Send('^{a}')
	Send($q3old)
	Send('{tab 6}')
	Send($art4min)
	Send('{tab}')
	Send($art4max)
	Send('{tab 2}')
	Send($plant)
	Send("{DOWN 11}")
	Send($proc)
	Send("{DOWN 18}")
	Send("{TAB 2}")
	send($sDatePast)
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
	Sleep(3000)
	Send('^{c}')
	$audiOldOrders = ClipGet()
	$audiOldOrders=StringRegExpReplace($audiOldOrders,'(\d)\s','\1')
	Sleep(3000)
	;~ COOIS skoda old orders
	ControlClick ('','','[ID:1001]','',1)
	ControlSend('','','[ID:1001]',$ncoois)
	Sleep(5000)
	WinWait( $title2, "", 0 )
	While 1
	   WinActivate( $title2, "" )
		  If WinExists( $title2 ) Then ExitLoop
	WEnd
	_ClipBoard_SetData($vData)
	Sleep(1000)
	Send('{tab}')
	Send('^{a}')
	Send($sk383old)
	Send('{tab 6}')
	Send($art4min)
	Send('{tab}')
	Send($art4max)
	Send('{tab 2}')
	Send($plant)
	Send("{DOWN 11}")
	Send($proc)
	Send("{DOWN 18}")
	Send("{TAB 2}")
	send($sDatePast)
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
	Sleep(3000)
	Send('^{c}')
	$skodaOldOrders = ClipGet()
	$skodaOldOrders=StringRegExpReplace($skodaOldOrders,'(\d)\s','\1')
	Sleep(3000)
	ControlClick ('','','[ID:1001]','',1)
	ControlSend('','','[ID:1001]',$n)
	Sleep(3000)
EndFunc


;~  SPLASH SCREEN WITH DATA SHOW

Func _ShowOldOrders()

	Global $mainWindow = GUICreate("Statistic monitor 0313", 700, 700, -1, -1, $WS_OVERLAPPEDWINDOW, $WS_EX_TOPMOST)

	Global $sFont = "Arial"
	GUISetFont(50,  $FW_NORMAL, $GUI_FONTUNDER, $sFont)
	GUICtrlCreateLabel('Старі замовлення / Old orders', 10, 10, 680, 50, BitOR($ES_AUTOVSCROLL, $SS_CENTER))

	GUISetFont(30,  $FW_NORMAL, $GUI_FONTUNDER, $sFont)
	Global $idLabelAudiFon = GUICtrlCreateLabel('Audi', 10, 70, 680, 120, BitOR($ES_AUTOVSCROLL, $SS_CENTER))

	GUISetFont(110,  $FW_NORMAL, $GUI_FONTITALIC, $sFont)
	Global $idLabelAudi = GUICtrlCreateLabel($audiOldOrders, 10, 150, 680, 180, BitOR($ES_AUTOVSCROLL, $SS_CENTER))

	GUISetFont(30,  $FW_NORMAL, $GUI_FONTUNDER, $sFont)
	Global $idLabelSkodaFon = GUICtrlCreateLabel('Skoda', 10, 330, 680, 100, BitOR($ES_AUTOVSCROLL, $SS_CENTER))

	GUISetFont(110,  $FW_NORMAL, $GUI_FONTITALIC, $sFont)
	Global $idLabelSkoda = GUICtrlCreateLabel($skodaOldOrders, 10, 410, 680, 165, BitOR($ES_AUTOVSCROLL, $SS_CENTER))

	GUISetFont(25,  $FW_NORMAL, $GUI_FONTUNDER, $sFont)
	Global $idDescri = GUICtrlCreateLabel('Ціль залишку не закритих замовлень, шт./день в розкрійному цеху', 10, 577, 680, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
	GUICtrlSetBkColor($idDescri, 0x00428df5  );tut

	GUISetFont(20,  $FW_NORMAL, $GUI_FONTITALIC, $sFont)
	Global $idGreenAudi = GUICtrlCreateLabel('Audi < 500', 10, 610, 220, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
	GUICtrlSetBkColor($idGreenAudi, 0x0011DD15)

	GUISetFont(20,  $FW_NORMAL, $GUI_FONTITALIC, $sFont)
	Global $idYellowAudi = GUICtrlCreateLabel('Audi 500 - 700', 240, 610, 220, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
	GUICtrlSetBkColor($idYellowAudi, $COLOR_YELLOW)

	GUISetFont(20,  $FW_NORMAL, $GUI_FONTITALIC, $sFont)
	Global $idRedAudi = GUICtrlCreateLabel('Audi > 700', 470, 610, 220, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
	GUICtrlSetBkColor($idRedAudi, $COLOR_RED)

	GUISetFont(20,  $FW_NORMAL, $GUI_FONTITALIC, $sFont)
	Global $idGreenSk = GUICtrlCreateLabel('Skoda < 200', 10, 660, 220, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
	GUICtrlSetBkColor($idGreenSk, 0x0011DD15)

	GUISetFont(20,  $FW_NORMAL, $GUI_FONTITALIC, $sFont)
	Global $idYellowSk = GUICtrlCreateLabel('Skoda 200 - 300', 240, 660, 220, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
	GUICtrlSetBkColor($idYellowSk, $COLOR_YELLOW)

	GUISetFont(20,  $FW_NORMAL, $GUI_FONTITALIC, $sFont)
	Global $idRedSk = GUICtrlCreateLabel('Skoda > 300', 470, 660, 220, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
	GUICtrlSetBkColor($idRedSk, $COLOR_RED)

	GUISetState(@SW_MAXIMIZE);@SW_MAXIMIZE

	If $skodaOldOrders < 200 Then
		  GUICtrlSetBkColor($idLabelSkoda, 0x0011DD15)
		  GUICtrlSetBkColor($idLabelSkodaFon, 0x0011DD15)
	   ElseIf $skodaOldOrders >= 200 And $skodaOldOrders <= 300 Then
		  GUICtrlSetBkColor($idLabelSkoda, $COLOR_YELLOW)
		  GUICtrlSetBkColor($idLabelSkodaFon, $COLOR_YELLOW)
	   ElseIf $skodaOldOrders > 300 Then
		  GUICtrlSetBkColor($idLabelSkoda, $COLOR_RED)
		  GUICtrlSetBkColor($idLabelSkodaFon, $COLOR_RED)
	   EndIf

	If $audiOldOrders < 500 Then
		  GUICtrlSetBkColor($idLabelAudi, 0x0011DD15)
		  GUICtrlSetBkColor($idLabelAudiFon, 0x0011DD15)
	   ElseIf $audiOldOrders >= 500 And $audiOldOrders <= 700 Then
		  GUICtrlSetBkColor($idLabelAudi, $COLOR_YELLOW)
		  GUICtrlSetBkColor($idLabelAudiFon, $COLOR_YELLOW)
	   ElseIf $audiOldOrders > 700 Then
		  GUICtrlSetBkColor($idLabelAudi, $COLOR_RED)
		  GUICtrlSetBkColor($idLabelAudiFon, $COLOR_RED)
	   EndIf
EndFunc

;~ STOCK ORDERS
Func _ShowStockOrders()

	Global $mainWindowStock = GUICreate("Statistic monitor 0313", 700, 700, -1, -1, $WS_OVERLAPPEDWINDOW, $WS_EX_TOPMOST)

	GUISetFont(50,  $FW_NORMAL, $GUI_FONTUNDER, $sFont)
	GUICtrlCreateLabel('Замовлення на складі / Stock level', 10, 10, 680, 50, BitOR($ES_AUTOVSCROLL, $SS_CENTER))

	GUISetFont(30,  $FW_NORMAL, $GUI_FONTUNDER, $sFont)
	Global $idLabelAudiFonStock = GUICtrlCreateLabel('Audi', 10, 70, 680, 120, BitOR($ES_AUTOVSCROLL, $SS_CENTER))

	GUISetFont(110,  $FW_NORMAL, $GUI_FONTITALIC, $sFont)
	Global $idLabelAudiStock = GUICtrlCreateLabel($audiStockOrders, 10, 150, 680, 180, BitOR($ES_AUTOVSCROLL, $SS_CENTER))

	GUISetFont(30,  $FW_NORMAL, $GUI_FONTUNDER, $sFont)
	Global $idLabelSkodaFonStock = GUICtrlCreateLabel('Skoda', 10, 330, 680, 100, BitOR($ES_AUTOVSCROLL, $SS_CENTER))

	GUISetFont(110,  $FW_NORMAL, $GUI_FONTITALIC, $sFont)
	Global $idLabelSkodaStock = GUICtrlCreateLabel($skodaStockOrders, 10, 410, 680, 165, BitOR($ES_AUTOVSCROLL, $SS_CENTER))

	GUISetFont(25,  $FW_NORMAL, $GUI_FONTUNDER, $sFont)
	Global $idDescriStock = GUICtrlCreateLabel('Ціль замовлень на складі швейного цеху, шт./день', 10, 577, 680, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
	GUICtrlSetBkColor($idDescriStock, 0x00428df5);tut

	GUISetFont(20,  $FW_NORMAL, $GUI_FONTITALIC, $sFont)
	Global $idGreenAudiStock = GUICtrlCreateLabel('Audi > 8700', 10, 610, 220, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
	GUICtrlSetBkColor($idGreenAudiStock, 0x0011DD15)

	GUISetFont(20,  $FW_NORMAL, $GUI_FONTITALIC, $sFont)
	Global $idYellowAudiStock = GUICtrlCreateLabel('Audi 8500 - 8700', 240, 610, 220, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
	GUICtrlSetBkColor($idYellowAudiStock, $COLOR_YELLOW)

	GUISetFont(20,  $FW_NORMAL, $GUI_FONTITALIC, $sFont)
	Global $idRedAudiStock = GUICtrlCreateLabel('Audi < 8500', 470, 610, 220, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
	GUICtrlSetBkColor($idRedAudiStock, $COLOR_RED)

	GUISetFont(20,  $FW_NORMAL, $GUI_FONTITALIC, $sFont)
	Global $idGreenSkStock = GUICtrlCreateLabel('Skoda > 2400', 10, 660, 220, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
	GUICtrlSetBkColor($idGreenSkStock, 0x0011DD15)

	GUISetFont(20,  $FW_NORMAL, $GUI_FONTITALIC, $sFont)
	Global $idYellowSkStock = GUICtrlCreateLabel('Skoda 2200 - 2400', 240, 660, 220, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
	GUICtrlSetBkColor($idYellowSkStock, $COLOR_YELLOW)

	GUISetFont(20,  $FW_NORMAL, $GUI_FONTITALIC, $sFont)
	Global $idRedSkStock = GUICtrlCreateLabel('Skoda < 2200', 470, 660, 220, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
	GUICtrlSetBkColor($idRedSkStock, $COLOR_RED)

	GUISetState(@SW_MAXIMIZE);@SW_MAXIMIZE

	If $skodaStockOrders > 2400 Then
		  GUICtrlSetBkColor($idLabelSkodaStock, 0x0011DD15)
		  GUICtrlSetBkColor($idLabelSkodaFonStock, 0x0011DD15)
	   ElseIf $skodaStockOrders >= 2200 And $skodaStockOrders <= 2400 Then
		  GUICtrlSetBkColor($idLabelSkodaStock, $COLOR_YELLOW)
		  GUICtrlSetBkColor($idLabelSkodaFonStock, $COLOR_YELLOW)
	   ElseIf $skodaStockOrders < 2200 Then
		  GUICtrlSetBkColor($idLabelSkodaStock, $COLOR_RED)
		  GUICtrlSetBkColor($idLabelSkodaFonStock, $COLOR_RED)
	   EndIf

	If $audiStockOrders > 8700 Then
		  GUICtrlSetBkColor($idLabelAudiStock, 0x0011DD15)
		  GUICtrlSetBkColor($idLabelAudiFonStock, 0x0011DD15)
	   ElseIf $audiStockOrders >= 8500 And $audiStockOrders <= 8700 Then
		  GUICtrlSetBkColor($idLabelAudiStock, $COLOR_YELLOW)
		  GUICtrlSetBkColor($idLabelAudiFonStock, $COLOR_YELLOW)
	   ElseIf $audiStockOrders < 8500 Then
		  GUICtrlSetBkColor($idLabelAudiStock, $COLOR_RED)
		  GUICtrlSetBkColor($idLabelAudiFonStock, $COLOR_RED)
	   EndIf
EndFunc


;~ STOCK ORDERS
Func _ShowDeliveredOrders()

	Global $mainWindowStock = GUICreate("Statistic monitor 0313", 700, 700, -1, -1, $WS_OVERLAPPEDWINDOW, $WS_EX_TOPMOST)
	
	GUISetFont(50,  $FW_NORMAL, $GUI_FONTUNDER, $sFont)
	GUICtrlCreateLabel('Доставлені замовлення / Dedicated orders', 10, 10, 680, 50, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
	
	
EndFunc
