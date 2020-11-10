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

HotKeySet("^{q}", "_quit")

Func _Date($iReturnTime = 1)
    Return @MDAY & '.' & @MON & '.' & @YEAR
EndFunc

Global $sFont = "Arial"
Global $title0 = "SAP Easy Access"
Global $title1 = "Material Document List"

Global $Q3N1MFB = "0"
Global $Q3FB = "0"
Global $Q3N1HFB = "0" 
Global $Q3N1MFC = "0"
Global $Q3FC = "0"
Global $Q3RB40 = "0"
Global $Q3RB60 = "0"
Global $Q3RC40 = "0"
Global $Q3RC60 = "0"

$nmb51 = "/nmb51{enter}"

Func _SetColor($data, $min, $max, $labelID)
	If $data > $max Then
        GUICtrlSetBkColor($labelID, 0x0011DD15)
     ElseIf $data >= $min And $data <= $max Then
        GUICtrlSetBkColor($labelID, $COLOR_YELLOW)
     ElseIf $data < $min Then
        GUICtrlSetBkColor($labelID, $COLOR_RED)
     EndIf
EndFunc

Func _ShowDeliveredOrders()
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
    $Q3RB60l = GUICtrlCreateLabel("Q3 RB 60",444,120,95,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1, 30, 700, 0, $sFont)
    GUICtrlSetBkColor(-1,"-2")
    $Q3RB60Dl = GUICtrlCreateLabel($Q3RB60,539,120,110,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
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
    $Q3RC40l = GUICtrlCreateLabel("Q3 RC 40",444,210,95,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1, 30, 700, 0, $sFont)
    GUICtrlSetBkColor(-1,"-2")
    $Q3RC40Dl = GUICtrlCreateLabel($Q3RC40,539,210,110,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1,50,700,0,"MS Sans Serif")
    GUICtrlSetBkColor(-1,"-2")
    $Q3RB40l = GUICtrlCreateLabel("Q3 RB 40",235,294,95,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1, 30, 700, 0, $sFont)
    GUICtrlSetBkColor(-1,"-2")
    $Q3RB40Dl = GUICtrlCreateLabel($Q3RB40,330,294,110,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1,50,700,0,"MS Sans Serif")
    GUICtrlSetBkColor(-1,"-2")
    $Q3RC60l = GUICtrlCreateLabel("Q3 RC 60",444,294,95,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
    GUICtrlSetFont(-1, 30, 700, 0, $sFont)
    GUICtrlSetBkColor(-1,"-2")
    $Q3RC60Dl = GUICtrlCreateLabel($Q3RC60,539,294,110,70,BitOr($SS_CENTER,$SS_NOTIFY),-1)
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
    GUICtrlSetBkColor(-1,"-2")
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

Func _delDelivered()
    GUIDelete($winDelivered)
EndFunc

_start()

Func _start()
    _ShowDeliveredOrders()
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
   _chooseData()
   _delDelivered()
   _show()
EndFunc

Func _show()
    $hTimer = TimerInit()
    While TimerDiff($hTimer)<1*60*1000
        _ShowDeliveredOrders()
       Sleep(30000)
       _delDelivered()
    WEnd
       Sleep(500)
       _start()
 EndFunc

Func _GetDataDeliver($point)
    $vData = "0"
    _ClipBoard_SetData($vData)
    Send("4000000")
    Send("{TAB}")
    Send("4999999")
    Send("{TAB 2}")
    Send("^a")
    Send("0313")
    Send("{DOWN}")
    Send("^a")
    Send("0313")
    Send("{DOWN 4}")
    Send("^a")
    Send("101")
    Send("{DOWN 5}")
    Send("^a")
    Send(_Date())
    Send("{DOWN 4}")
    Send("^a")
    Send($point)
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
    ControlSend($title1, "", "[CLASS:Edit; INSTANCE:1; ID:1001]", $nmb51 & "{ENTER}")
	While 1
	   WinActivate($title1)
	      If ControlGetFocus($title1) <> "SAPALVGrid1" Then ExitLoop
    WEnd
    Return $point
EndFunc    

Func _chooseData()
	WinWait( $title0, "", 0 )
	If WinExists( $title0 ) Then
		WinActivate( $title0 )
	EndIf
	   WinSetState($title0,"",@SW_MAXIMIZE)
	   ControlSend($title0, "", "[CLASS:Edit; INSTANCE:1; ID:1001]", $nmb51 & "{ENTER}")
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
    ControlSend($title1, "", "[CLASS:Edit; INSTANCE:1; ID:1001]", "/n" & "{ENTER}")
EndFunc

Func _quit()
	Exit
EndFunc