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

AutoItSetOption ( "SendKeyDelay", 100 )
AutoItSetOption ( "SendKeyDownDelay", 100 )
AutoItSetOption ( "WinSearchChildren", 1 )
AutoItSetOption ( "WinTitleMatchMode", 4 )
AutoItSetOption ( "WinWaitDelay", 500 )

HotKeySet("q", "_quit")

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

;~ Show window with delivered orders
Func _ShowDeliveredOrdersAudi()
    Global $winDelivered = GUICreate("Statistic monitor 0313",700,700,-1,-1,-1,BitOr($WS_EX_TOPMOST,$WS_EX_OVERLAPPEDWINDOW))
    GUICtrlCreateLabel("Доставлені замовлення / Dedicated orders",0,10,700,70,$SS_CENTER,-1)
    GUICtrlSetFont(-1,50,400,0,"MS Sans Serif")
    GUICtrlSetBkColor(-1,"-2")
    $Q3FBNORMl = GUICtrlCreateLabel("Q3 FB Norm",20,120,95,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1, 30, 700, 0, $sFont)
    GUICtrlSetBkColor(-1,"-2")
    $Q3FBNORMDl = GUICtrlCreateLabel($Q3N1MFB,115,120,110,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1,50,700,0,"MS Sans Serif")
    GUICtrlSetBkColor(-1,"-2")
    $Q3FCNORMl = GUICtrlCreateLabel("Q3 FC Norm",235,120,95,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1, 30, 700, 0, $sFont)
    GUICtrlSetBkColor(-1,"-2")
    $Q3FCNORMDl = GUICtrlCreateLabel($Q3N1MFC,330,120,110,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1,50,700,0,"MS Sans Serif")
    GUICtrlSetBkColor(-1,"-2")
    $Q3RB60l = GUICtrlCreateLabel("Q3 RB 60",450,120,95,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1, 30, 700, 0, $sFont)
    GUICtrlSetBkColor(-1,"-2")
    $Q3RB60Dl = GUICtrlCreateLabel($Q3RB60,545,120,110,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1,50,700,0,"MS Sans Serif")
    GUICtrlSetBkColor(-1,"-2")
    $Q3FBSPl = GUICtrlCreateLabel("Q3 FB Sp",20,210,95,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1, 30, 700, 0, $sFont)
    GUICtrlSetBkColor(-1,"-2")
    $Q3FBSPDl = GUICtrlCreateLabel($Q3FB,115,210,110,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1,50,700,0,"MS Sans Serif")
    GUICtrlSetBkColor(-1,"-2")
    $Q3FCSPl = GUICtrlCreateLabel("Q3 FC Sp",235,210,95,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1, 30, 700, 0, $sFont)
    GUICtrlSetBkColor(-1,"-2")
    $Q3FCSPDl = GUICtrlCreateLabel($Q3FC,330,210,110,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1,50,700,0,"MS Sans Serif")
    GUICtrlSetBkColor(-1,"-2")
    $Q3RC40l = GUICtrlCreateLabel("Q3 RC 40",450,210,95,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1, 30, 700, 0, $sFont)
    GUICtrlSetBkColor(-1,"-2")
    $Q3RC40Dl = GUICtrlCreateLabel($Q3RC40,545,210,110,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1,50,700,0,"MS Sans Serif")
    GUICtrlSetBkColor(-1,"-2")
    $Q3RB40l = GUICtrlCreateLabel("Q3 RB 40",235,294,95,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1, 30, 700, 0, $sFont)
    GUICtrlSetBkColor(-1,"-2")
    $Q3RB40Dl = GUICtrlCreateLabel($Q3RB40,330,294,110,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1,50,700,0,"MS Sans Serif")
    GUICtrlSetBkColor(-1,"-2")
    $Q3RC60l = GUICtrlCreateLabel("Q3 RC 60",450,294,95,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1, 30, 700, 0, $sFont)
    GUICtrlSetBkColor(-1,"-2")
    $Q3RC60Dl = GUICtrlCreateLabel($Q3RC60,545,294,110,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1,50,700,0,"MS Sans Serif")
    GUICtrlSetBkColor(-1,"-2")
    $Q3FBSSPDl = GUICtrlCreateLabel($Q3N1HFB,115,294,110,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1,50,700,0,"MS Sans Serif")
    GUICtrlSetBkColor(-1,"-2")
    $Q3FBSSPl = GUICtrlCreateLabel("Q3 FB Ssp",20,294,95,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1, 30, 700, 0, $sFont)
    GUICtrlSetBkColor(-1,"-2")
    $target_for_audi = GUICtrlCreateLabel("Ціль Audi на складі швейного цеху по компонентах",0,420,700,30,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1,30,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($target_for_audi, 0x00428df5)
    ;~ Static data for description (create func ?)
    $Q3FBNORMlt = GUICtrlCreateLabel('Q3 FB Norm', 20, 480, 120, 30, BitOR($ES_AUTOVSCROLL, $SS_LEFT))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
	$Q3FBNORMlt = GUICtrlCreateLabel('<170', 100, 480, 40, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_RED)
	$Q3FBNORMlt = GUICtrlCreateLabel('170-570', 137, 480, 54, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_YELLOW)
    $Q3FBNORMlt = GUICtrlCreateLabel('>570', 190, 480, 35, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, 0x0011DD15)
    $Q3FBNORMlt = GUICtrlCreateLabel('Q3 FB Sp', 20, 520, 120, 30, BitOR($ES_AUTOVSCROLL, $SS_LEFT))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
	$Q3FBNORMlt = GUICtrlCreateLabel('<180', 100, 520, 40, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_RED)
	$Q3FBNORMlt = GUICtrlCreateLabel('180-600', 137, 520, 54, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_YELLOW)
    $Q3FBNORMlt = GUICtrlCreateLabel('>600', 190, 520, 35, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, 0x0011DD15)    
    $Q3FBNORMlt = GUICtrlCreateLabel('Q3 FB Ssp', 20, 560, 120, 30, BitOR($ES_AUTOVSCROLL, $SS_LEFT))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
	$Q3FBNORMlt = GUICtrlCreateLabel('<30', 100, 560, 40, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_RED)
	$Q3FBNORMlt = GUICtrlCreateLabel('30-90', 137, 560, 54, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_YELLOW)
    $Q3FBNORMlt = GUICtrlCreateLabel('>90', 190, 560, 35, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, 0x0011DD15)
    $Q3FBNORMlt = GUICtrlCreateLabel('Q3 FC Norm', 235, 480, 120, 30, BitOR($ES_AUTOVSCROLL, $SS_LEFT))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
	$Q3FBNORMlt = GUICtrlCreateLabel('<160', 315, 480, 40, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_RED)
    $Q3FBNORMlt = GUICtrlCreateLabel('160-550', 352, 480, 54, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_YELLOW)
    $Q3FBNORMlt = GUICtrlCreateLabel('>550', 405, 480, 35, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, 0x0011DD15)
    $Q3FBNORMlt = GUICtrlCreateLabel('Q3 FC Sp', 235, 520, 120, 30, BitOR($ES_AUTOVSCROLL, $SS_LEFT))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
	$Q3FBNORMlt = GUICtrlCreateLabel('<220', 315, 520, 40, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_RED)
    $Q3FBNORMlt = GUICtrlCreateLabel('220-750', 352, 520, 54, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_YELLOW)
    $Q3FBNORMlt = GUICtrlCreateLabel('>750', 405, 520, 35, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, 0x0011DD15)
    $Q3FBNORMlt = GUICtrlCreateLabel('Q3 RB 40', 235, 560, 120, 30, BitOR($ES_AUTOVSCROLL, $SS_LEFT))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
	$Q3FBNORMlt = GUICtrlCreateLabel('<220', 315, 560, 40, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_RED)
    $Q3FBNORMlt = GUICtrlCreateLabel('220-750', 352, 560, 54, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_YELLOW)
    $Q3FBNORMlt = GUICtrlCreateLabel('>750', 405, 560, 35, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, 0x0011DD15)
    $Q3FBNORMlt = GUICtrlCreateLabel('Q3 RB 60', 450, 480, 120, 30, BitOR($ES_AUTOVSCROLL, $SS_LEFT))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    $Q3FBNORMlt = GUICtrlCreateLabel('<220', 525, 480, 40, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_RED)
    $Q3FBNORMlt = GUICtrlCreateLabel('220-750', 562, 480, 54, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_YELLOW)
    $Q3FBNORMlt = GUICtrlCreateLabel('>750', 615, 480, 35, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, 0x0011DD15)
    $Q3FBNORMlt = GUICtrlCreateLabel('Q3 RC 40', 450, 520, 120, 30, BitOR($ES_AUTOVSCROLL, $SS_LEFT))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    $Q3FBNORMlt = GUICtrlCreateLabel('<220', 525, 520, 40, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_RED)
    $Q3FBNORMlt = GUICtrlCreateLabel('220-750', 562, 520, 54, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_YELLOW)
    $Q3FBNORMlt = GUICtrlCreateLabel('>750', 615, 520, 35, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, 0x0011DD15)
    $Q3FBNORMlt = GUICtrlCreateLabel('Q3 RC 60', 450, 560, 120, 30, BitOR($ES_AUTOVSCROLL, $SS_LEFT))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    $Q3FBNORMlt = GUICtrlCreateLabel('<220', 525, 560, 40, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_RED)
    $Q3FBNORMlt = GUICtrlCreateLabel('220-750', 562, 560, 54, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_YELLOW)
    $Q3FBNORMlt = GUICtrlCreateLabel('>750', 615, 560, 35, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, 0x0011DD15)

    GUISetState(@SW_SHOW,$winDelivered)
    GUISetState(@SW_MAXIMIZE)

    _SetColor($Q3N1MFB, 170, 570, $Q3FBNORMDl)
    _SetColor($Q3N1MFB, 170, 570, $Q3FBNORMl)

    _SetColor($Q3FB, 180, 600, $Q3FBSPDl)
    _SetColor($Q3FB, 180, 600, $Q3FBSPl)

    _SetColor($Q3N1HFB, 30, 90, $Q3FBSSPDl)
    _SetColor($Q3N1HFB, 30, 90, $Q3FBSSPl)

    _SetColor($Q3N1MFC, 160, 550, $Q3FCNORMDl)
    _SetColor($Q3N1MFC, 160, 550, $Q3FCNORMl)

    _SetColor($Q3FC, 220, 750, $Q3FCSPDl)
    _SetColor($Q3FC, 220, 750, $Q3FCSPl)

    _SetColor($Q3RB40, 220, 750, $Q3RB40Dl)
    _SetColor($Q3RB40, 220, 750, $Q3RB40l)
    
    _SetColor($Q3RB60, 220, 750, $Q3RB60Dl)
    _SetColor($Q3RB60, 220, 750, $Q3RB60l)

    _SetColor($Q3RC40, 220, 750, $Q3RC40Dl)
    _SetColor($Q3RC40, 220, 750, $Q3RC40l)
    
    _SetColor($Q3RC60, 220, 750, $Q3RC60Dl)
    _SetColor($Q3RC60, 220, 750, $Q3RC60l)
EndFunc

Func _delDeliveredAudi()
    GUIDelete($winDelivered)
EndFunc



Func _ShowDeliveredOrdersSkoda()
    Global $winDeliveredSkoda = GUICreate("Statistic monitor 0313",700,700,-1,-1,-1,BitOr($WS_EX_TOPMOST,$WS_EX_OVERLAPPEDWINDOW))
    GUICtrlCreateLabel("Доставлені замовлення / Dedicated orders",0,10,700,70,$SS_CENTER,-1)
    GUICtrlSetFont(-1,50,400,0,"MS Sans Serif")
    GUICtrlSetBkColor(-1,"-2")
    $SKRB40l = GUICtrlCreateLabel("SK RB 40",20,120,95,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1, 30, 700, 0, $sFont)
    GUICtrlSetBkColor(-1,"-2")
    $SKRB40Dl = GUICtrlCreateLabel($SKRB40,115,120,110,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1,50,700,0,"MS Sans Serif")
    GUICtrlSetBkColor(-1,"-2")
    $SKRB60l = GUICtrlCreateLabel("SK RB 60",235,120,95,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1, 30, 700, 0, $sFont)
    GUICtrlSetBkColor(-1,"-2")
    $SKRB60Dl = GUICtrlCreateLabel($SKRB60,330,120,110,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1,50,700,0,"MS Sans Serif")
    GUICtrlSetBkColor(-1,"-2")
    $SKRCl = GUICtrlCreateLabel("SK RC",450,120,95,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1, 30, 700, 0, $sFont)
    GUICtrlSetBkColor(-1,"-2")
    $SKRCDl = GUICtrlCreateLabel($SKRC,545,120,110,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1,50,700,0,"MS Sans Serif")
    GUICtrlSetBkColor(-1,"-2")
    $target_for_skoda = GUICtrlCreateLabel("Ціль Skoda на складі швейного цеху по компонентах",0,420,700,30,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1,30,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($target_for_skoda, 0x00428df5)
    ;~ Static data for description (create func ?) !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    $Q3FBNORMlt = GUICtrlCreateLabel('SK RB 40', 20, 480, 120, 30, BitOR($ES_AUTOVSCROLL, $SS_LEFT))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    $Q3FBNORMlt = GUICtrlCreateLabel('<105', 100, 480, 40, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_RED)
    $Q3FBNORMlt = GUICtrlCreateLabel('150-350', 137, 480, 54, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_YELLOW)
    $Q3FBNORMlt = GUICtrlCreateLabel('>350', 190, 480, 35, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, 0x0011DD15)
    $Q3FBNORMlt = GUICtrlCreateLabel('SK RB 60', 235, 480, 120, 30, BitOR($ES_AUTOVSCROLL, $SS_LEFT))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    $Q3FBNORMlt = GUICtrlCreateLabel('<105', 315, 480, 40, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_RED)
    $Q3FBNORMlt = GUICtrlCreateLabel('150-350', 352, 480, 54, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_YELLOW)
    $Q3FBNORMlt = GUICtrlCreateLabel('>350', 405, 480, 35, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, 0x0011DD15)
    $Q3FBNORMlt = GUICtrlCreateLabel('SK RC ', 450, 480, 120, 30, BitOR($ES_AUTOVSCROLL, $SS_LEFT))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    $Q3FBNORMlt = GUICtrlCreateLabel('<105', 525, 480, 40, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_RED)
    $Q3FBNORMlt = GUICtrlCreateLabel('105-350', 562, 480, 54, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, $COLOR_YELLOW)
    $Q3FBNORMlt = GUICtrlCreateLabel('>350', 615, 480, 35, 30, BitOR($ES_AUTOVSCROLL, $SS_CENTER))
    GUICtrlSetFont(-1,27,700,0,"MS Sans Serif")
    GUICtrlSetBkColor($Q3FBNORMlt, 0x0011DD15)
    
    GUISetState(@SW_SHOW,$winDeliveredSkoda)
    GUISetState(@SW_MAXIMIZE)

    _SetColor($SKRB40, 105, 350, $SKRB40Dl)
    _SetColor($SKRB40, 105, 350, $SKRB40l)

    _SetColor($SKRB60, 105, 350, $SKRB60Dl)
    _SetColor($SKRB60, 105, 350, $SKRB60l)

    _SetColor($SKRC, 105, 350, $SKRCDl)
    _SetColor($SKRC, 105, 350, $SKRCl)

EndFunc

Func _delDeliveredSkoda()
    GUIDelete($winDeliveredSkoda)
EndFunc

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

    GUISetState(@SW_SHOW,$mainWindow)
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

Func _delOld()
    GUIDelete($mainWindow)
EndFunc

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
    
    GUISetState(@SW_SHOW,$mainWindowStock)
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

Func _delStock()
    GUIDelete($mainWindowStock)
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
    _ClipBoard_SetData($artMin4)
    Send('+{INS}')
    Send("{TAB}")
    _ClipBoard_SetData($artMax4)
    Send('+{INS}')
    Send("{TAB 2}")
    Send("^a")
    _ClipBoard_SetData($cutPlant)
    Send('+{INS}')
    Send("{DOWN}")
    Send("^a")
    _ClipBoard_SetData($cutPlant)
    Send('+{INS}')
    Send("{DOWN 4}")
    Send("^a")
    _ClipBoard_SetData($movement)
    Send('+{INS}')
    Send("{DOWN 2}")
    Send("^a")
    Send("{DEL}")
    Send("{DOWN 3}")
    Send("^a")
    _ClipBoard_SetData(_Date())
    Send('+{INS}')
    Send("{DOWN 4}")
    Send("^a")
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
    Send("^{a}")
    _ClipBoard_SetData($artMin4)
    Send('+{INS}')
	Send("{TAB}")
    _ClipBoard_SetData($artMax4)
    Send('+{INS}')
	Send("{TAB 2}")
    _ClipBoard_SetData($sewPlant)
    Send('+{INS}')
	Send("{DOWN}")
	Send("^{a}")
	Send("{DEL}")
	Send("{DOWN 9}")
	Send("+{TAB}")
	Send("{DOWN}")
	Send("{TAB}")
	Send("^{a}")
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
    Send('^{a}')
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