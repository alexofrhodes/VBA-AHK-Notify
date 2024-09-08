#Requires Autohotkey v2.0-

/* 2024 09
Modified the GenCallback to be called from vba to run macros on notification click
 https://github.com/alexofrhodes
*/


/* The Automator
* ============================================================================ *
* Want a clear path for learning AutoHotkey?                                   *
* Take a look at our AutoHotkey courses here: the-Automator.com/Discover          *
* They're structured in a way to make learning AHK EASY                        *
* And come with a 200% moneyback guarantee so you have NOTHING to risk!        *
* ============================================================================ *
*/

/*
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	;;;;;;;;;;;;;;;Notify AHK V2;;;;;;;;;;;;;;;;;
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	Header, Body, and Background colors supported: Black,Silver,Gray,White,Maroon,Red,Purple,Fuchsia,Green,Lime,Olive,Yellow,Navy,Blue,Teal,Aqua

	'Duration' is how long you want the notice to be displayed before it disappears.
	It should be number, Use 0 to leave it on screen indefintiely until user clicks.

	Be sure to Include this library when using from another script.
	Example:  #Include <Notifyv2>

*/

For arg in A_ARGS
{
	if A_Index &1
	{
		; msgbox arg "`n" A_ARGS[A_Index + 1]
		switch arg {
		case '-BDText':
			msg := A_ARGS[A_Index + 1]
		Case '-Link':
			;msgbox arg "`n" A_ARGS[A_Index + 1]
			href := A_ARGS[A_Index + 1]

			if !A_IsCompiled
			&& href ~= '"'
				href := RegExReplace(href,'\"+','"')
			else if !A_IsCompiled
			&& href ~= '"' == false
				href := RegExReplace(href,'<a href=(.*?)>','<a href="$1">')

			msg := {link: href}
		case '-GenIcon', '-HDFontSize', '-BDFontSize','-GenDuration','-GenIconSize':
			Notify.Default.%StrReplace(Trim(arg), '-')% := number(A_ARGS[A_Index + 1])
		case '-HDText', '-HDFontColor','-HDFont','-BDText','-BDFontColor', '-BDFont','-GenBGColor','-GenSound':
			Notify.Default.%StrReplace(Trim(arg), '-')% := A_ARGS[A_Index + 1]
		case '-GenCallback':  
			Notify.Default.%StrReplace(Trim(arg), '-')% := A_ARGS[A_Index + 1]
		default:
		}
	}
}
if IsSet(msg)
{
	Notify.show(msg)
}

Class Notify
{
	; do not modify this variable directly because
	; it would cause the main script to break
	static _Default := {
		HDText        : "",
		HDFontSize    : 16,
		HDFont        : "Impact",
		HDFontColor   : "0x298939",
		BDText        : "Click to Callback",
		BDFontSize    : 12,
		BDFontColor   : "Black",
		BDFont        : "Book Antiqua",
		GenBGColor    : "0xFFD23E",
		GenDuration   : 3,
		GenSound      : "",
		GenIcon	      : "",
		GenIconSize   : 30,
		GenMonitor    : MonitorGetPrimary(),
		GenLoc        : 'RightBottom',
		GenCallback	  : "",
	}

	Static Default
	{
		set => Notify._default := value
		get {
			static default_props := Map(
			'HDText'        , "",
			'HDFontSize'    , 14,
			'HDFont'        , "Impact",
			'HDFontColor'   , "Black",
			'BDFontSize'    , 10,
			'BDFontColor'   , "0x298939",
			'BDFont'        , "Book Antiqua",
			'GenBGColor'    , "0xFFD23E",
			'GenDuration'   , 3,
			'GenSound'      , "",
			'GenIcon'       , "",
			'GenIconSize'   , 30,
			'GenMonitor'    , MonitorGetPrimary(),
			'GenLoc'        , 'RightBottom',
			'GenCallback'	, "",
			)
			for prop in default_props
				if !Notify._Default.HasProp(prop)
					Notify._Default.%prop% := default_props[prop]

			return Notify._default
		}
	}
	Static wavList := "Sound List:`nName" ; `t`t Path"
	Static wav := Notify.GetSoundFiles()

	static Show(Input)
	{

		Switch(Type(Input))
		{
		Case "String":
			this.HDText      := Notify.Default.HDText
			this.HDSize      := Notify.Default.HDFontSize
			this.HDColor     := Notify.Default.HDFontColor
			this.HDFont      := Notify.Default.HDFont
			this.Text        := input
			this.BDSize      := Notify.Default.BDFontSize
			this.BDColor     := Notify.Default.BDFontColor
			this.BDFont      := Notify.Default.BDFont
			this.Duration    := Notify.Default.GenDuration
			this.Color       := Notify.Default.GenBGColor
			this.Sound       := Notify.Default.GenSound
			this.GenIcon     := Notify.Default.GenIcon
			this.GenIconSize := Notify.Default.GenIconSize
			this.GenMonitor  := Notify.Default.GenMonitor
			this.GenLoc      := Notify.Default.GenLoc
			this.Link        := ""
			this.Callback    := Notify.Default.GenCallback
		Case "Object":
			this.HDText      := input.HasOwnProp("HDText")      ? input.HDText      : Notify.Default.HDText
			this.HDSize      := input.HasOwnProp("HDFontSize")  ? input.HDFontSize  : Notify.Default.HDFontSize
			this.HDColor     := input.HasOwnProp("HDFontColor") ? input.HDFontColor : Notify.Default.HDFontColor
			this.HDFont      := input.HasOwnProp("HDFont")      ? input.HDFont      : Notify.Default.HDFont
			this.BDSize      := input.HasOwnProp("BDFontSize")  ? input.BDFontSize  : Notify.Default.BDFontSize
			this.BDColor     := input.HasOwnProp("BDFontColor") ? input.BDFontColor : Notify.Default.BDFontColor
			this.BDFont      := input.HasOwnProp("BDFont")      ? input.BDFont      : Notify.Default.BDFont
			this.Color       := input.HasOwnProp("GenBGColor")  ? input.GenBGColor  : Notify.Default.GenBGColor
			this.Duration    := input.HasOwnProp("GenDuration") ? input.GenDuration : Notify.Default.GenDuration
			this.Sound       := input.HasOwnProp("GenSound")    ? input.GenSound    : Notify.Default.GenSound
			this.GenIcon     := input.HasOwnProp("GenIcon")     ? input.GenIcon     : Notify.Default.GenIcon
			this.GenIconSize := input.HasOwnProp("GenIconSize") ? input.GenIconSize : Notify.Default.GenIconSize
			this.GenMonitor  := input.HasOwnProp("GenMonitor")  ? input.GenMonitor  : Notify.Default.GenMonitor
			this.GenLoc      := input.HasOwnProp("GenLoc")      ? input.GenLoc      : Notify.Default.GenLoc
			this.Link        := input.HasOwnProp("Link")        ? input.Link        : ""
			this.Callback    := input.HasOwnProp("GenCallback") ? input.GenCallback : ""
			this.Text        := input.HasOwnProp("BDText")      ? input.BDText      : ""

		}
		Notify.Play(this.Sound)
		this.Notice := MultiGui(this)
		if this.Duration != 0
			this.Close()
		return this
	}

	Static CloseLast()
	{
		try Notify.Notice.Close()
	}

	Static Close()
	{
		fn := ObjBindMethod(this, "animation", this.Notice)
		Settimer fn, -(this.Duration * 1000)
	}

	static animation(Notice) => Notice.close()

	static Play(Sound)
	{
		if RegExMatch(Sound,'^\*\-?\d+')
		|| FileExist(Sound)
			return Soundplay(Sound)
		try SoundFile := Notify.wav[Sound]
		catch
			return
		if FileExist(SoundFile)
			Soundplay(SoundFile)
		return
	}

	Static GetSoundFiles()
	{
		wav := map()
		loop files, "C:\Windows\Media\*.wav"
		{
			name := RegExReplace(A_LoopFileName,"Windows |notify |Hardware |.wav")
			if InStr(name," ")
				continue
			this.wavList .= "`n"  name  ;(InStr(name,"Alarm") ? "`t" : StrLen(name) < 8 ? "`t`t":"`t" ) ": " A_LoopFileName
			wav[name] := A_LoopFileFullPath
		}

		loop files, A_ScriptDir "\res\*.wav"
		{
			name := StrReplace(A_LoopFileName,".wav")
			this.wavList .= "`n"  name ;(StrLen(name) < 8 ? "`t`t":"`t" ) ": " A_LoopFileName
			wav[name] := A_LoopFileFullPath
		}
		return wav
	}


	; method to list all supported Alert Sounds
	Static SoundList() => this.Show({HDText:"GenSound list`nSupported by Notify",BDText:'Copied to clipboard`n' A_Clipboard := this.wavList,GenDuration:0,GenSound:"Insert"})
	; method to List all Color
	Static ColorList()
	{
		Colors :="
		(
			Black
			Silver
			Gray
			White
			Maroon
			Red
			Purple
			Fuchsia
			Green
			Lime
			Olive
			Yellow
			Navy
			Blue
			Teal
			Aqua
		)"
		this.Show({HDText:"HD BD and Gen Colors`nSupported by Notify",BDText:'Copied to clipboard`n' A_Clipboard := Colors,GenDuration:0,GenSound:"Remove"})
	}
	; method to list all GenIcons
	Static GenIconList()
	{
		GenIconHelp :=
		(
			'GenIcon can be:
			• Integer from Shell32.dll
			• Image/Icon Path
			• Any of the following strings:
				o Critical
				o Question
				o Exclamation
				o Information
				o Security

			GenIconSize: is number where the hight and width are the same'
		)
		Notify.Show({
			HDText:"GenIcon List`nGenIcons number or address should be passed",
			BDFontSize:16,
			GenDuration:10,
			GenIcon:96,
			GenIconSize:50,
			BDText: GenIconHelp
		})
	}


	static IconPicker()
	{
		Count := 329, Shell := 1, Image := 0, icoFile := "shell32.dll", Height := A_ScreenHeight - 170 ;Define constants
		iGui := Gui('-MinimizeBox -MaximizeBox','Notify Icon Picker')
		iGui.OnEvent('Close',exit)
		LV := iGui.AddListView('h' Height ' w400 +Icon',['Number'])
		LV.OnEvent('click',ListClick)
		ImageListID := IL_Create(Count,10,true)
		LV.SetImageList(ImageListID)
		loop Count
		{
			pos := IL_Add(ImageListID,icoFile,A_Index)
			LV.Add("Icon" pos,A_index)
		}
		LV.ModifyCol(1,'autohdr')  ; Auto-adjust the column widths.
		LV.ModifyCol(2,'autohdr integer Center')  ; Auto-adjust the column widths.
		iGui.Show()
		return

		ListClick(obj,info){
			n := LV.getText(info )
			a_Clipboard := n ; "Menu, Tray, Icon, %A_WinDir%\system32\" IcoFile "," info " `;Set custom Script icon`n"
			tooltip 'Copied Icon Number ' n
			SetTimer( ToolTip, -800  )
		}

		exit(*)
		{
			iGui.Destroy()
		}
	}

	Static DisplayCheck()
	{
		MonitorCount := MonitorGetCount()
		MonitorPrimary := MonitorGetPrimary()
		Notify.show(
			{
				HDText:'Monitor Info',
				BDText: 'Monitor Count: ' MonitorCount '`nPrimary Monitor: ' MonitorPrimary '`nClick to close',
				GenDuration:0
			}
		)
		Loop MonitorCount
		{
			MonitorGet A_Index, &L, &T, &R, &B
			MonitorGetWorkArea A_Index, &WL, &WT, &WR, &WB
			Notify.show(
			{
				HDText:'Monitor:`t#' A_Index ,
				BDText: 
				(
				'Name:`t' MonitorGetName(A_Index) '
				Left:`t' L ' (' WL ' work)
				Top:`t' T ' (' WT ' work)
				Right:`t' R ' (' WR ' work)
				Bottom:`t' B ' (' WB ' work)
				`t`tClick to close'
				)
				,
				GenLoc: 'C',
				GenMonitor:a_index,
				GenDuration:0
			}
			)
		} 
	}

}

Class MultiGui
{
	static Guis := array()
	; Static Taskbar := MultiGui.GetTaskBarPos()
	Static Monitors := MultiGui.CalcMonitor()
	Static LastPOs := map() ;{x:0,y:0,w:0,h:0}
	Static ShellDll := A_WinDir "\System32\shell32.dll"
	Static user32Dll := A_WinDir "\system32\user32.dll"
	Static Warning := Map(
		"Exclamation",2,
		"Question",3,
		"Critical",4,
		"Information",5,
		"Security",7
	)
	__new(info)
	{
		if info.GenMonitor > MultiGui.Monitors.length 
			info.GenMonitor := MonitorGetPrimary()

		MyGui := Gui("-Caption +AlwaysOnTop +Owner +LastFound")
		MyGui.MarginX := 5
		MyGui.MarginY := 5
		MyGui.BackColor := info.Color
		if (Type(Info.GenIcon) = "Integer")
			MyGui.AddPicture("w" Info.GenIconSize " h" Info.GenIconSize " Icon" Info.GenIcon + 0, MultiGui.ShellDll)
		else if FileExist(Info.GenIcon)
			MyGui.AddPicture("w" Info.GenIconSize " h" Info.GenIconSize,Info.GenIcon )
		else if Info.GenIcon && InStr("Critical,Question,Exclamation,Information,Security",Info.GenIcon)
			MyGui.AddPicture("w" Info.GenIconSize " h" Info.GenIconSize " Icon" MultiGui.Warning[Info.GenIcon], MultiGui.user32Dll)

		MyGui.SetFont("c" info.HDColor " s" info.HDSize , info.HDFont )
		if info.HDText
			MyGui.Add("Text","x+m", info.HDText)
		MyGui.SetFont(opts := "c" info.BDColor " s" info.BDSize , info.BDFont )
		if info.Link
			MyGui.AddLink(, info.Link)
		else if info.Text
			MyGui.AddText("y+m",info.Text ) ;"xp yp+" this.Header.Font.Size +9.5, this.Body.Text)
		MyGui.Show("Hide")
		This.MyGui := MyGui
		WinGetPos(&x,&y,&w,&h,MyGui)

		clickArea := MyGui.Add("Text", "x0 y0 w" . W . " h" . H . " BackgroundTrans")

		if !(info.Callback = "")
		{
			clickArea.OnEvent("Click", ObjBindMethod(this,"xlRun", info.Callback) )	
		    Info.Duration := 0
		}

		if Info.Duration = 0
			clickArea.OnEvent("Click", ObjBindMethod(this,"Close",MyGui) )
		MyGui.Monitors := info.GenMonitor
		MultiGui.Guis.Push(MyGui)



		if RegExMatch(info.GenLoc,'(?=.*?(?<x>x\d+))(?=.*?(?<y>y\d+))',&Out)
			POS := Out.x ' ' Out.y, LOC := 'xy'
		else
			POs := MultiGui.GeneratePOS(info,x,y,w,h,&Loc)
		
		MyGui.Show(POS " NoActivate")
		WinGetPos(&x,&y,&w,&h,MyGui)
		MultiGui.LastPOs[info.GenMonitor LOC] := {x:x,y:y,w:w,h:h}
	}

	static GeneratePOS(info,x,y,w,h,&Loc)
	{
		n := info.GenMonitor
		Switch info.GenLoc, 0
		{
			Case 'Center', 'C':
				Loc := 'C'
				
				; We are clculating the center position
				; to make sure that it displays correctly in all monitors
				; not only the primary monitor
				POS := "x" (MultiGui.Monitors[n]['x'] + MultiGui.Monitors[n]['w'])/2 - (w/2)  " y" (MultiGui.Monitors[n]['y'] + MultiGui.Monitors[n]['h'])/2 - (h/2)
			Case 'TopLeft','TL', 'LeftTop','LT':
				Loc := 'LT'
				if MultiGui.LastPOs.Has(n Loc)
					POS := "x" MultiGui.Monitors[n]['x'] " y" MultiGui.LastPOs[n Loc].y + MultiGui.LastPOs[n Loc].h +  1
				else
					POS := "x" MultiGui.Monitors[n]['x'] " y" MultiGui.Monitors[n]['y']
			Case 'TopRight','TR','RightTop','RT':
				Loc := 'RT'
				if MultiGui.LastPOs.Has(n Loc)
					POS := "x" MultiGui.Monitors[n]['w'] - w " y" MultiGui.LastPOs[n Loc].y + MultiGui.LastPOs[n Loc].h + 1
				else
					POS := "x" MultiGui.Monitors[n]['w'] - w " y" MultiGui.Monitors[n]['y'] 
			Case 'LeftBottom','LB','BottomLeft','BL':
				Loc := 'LB'
				if MultiGui.LastPOs.Has(n Loc)
					POS := "x" MultiGui.Monitors[n]['x'] " y" MultiGui.LastPOs[n Loc].y - h - 1
				else
					POS := "x" MultiGui.Monitors[n]['x'] " y" MultiGui.Monitors[n]['h'] - h
			Case 'RightBottom','RB','BottomRight','BR':
				Loc := 'BR'
				if MultiGui.LastPOs.Has(n Loc)
					POS := "x" MultiGui.Monitors[n]['w'] - w " y" MultiGui.LastPOs[n Loc].y - h - 1
				else
					POS := "x" MultiGui.Monitors[n]['w'] - w " y" MultiGui.Monitors[n]['h'] - h
			Default: ; default is right bottom
				Loc := 'BR'
				if MultiGui.LastPOs.Has(n Loc)
					POS := "x" MultiGui.Monitors[n]['w'] - w " y" MultiGui.LastPOs[n Loc].y - h - 1
				else
					POS := "x" MultiGui.Monitors[n]['w'] - w " y" MultiGui.Monitors[n]['h'] - h
		}
		return POS
	}

	xlRun(fullMacroName,*) {  
		; MsgBox(fullMacroName)

		xlApp := Excel_Get()
		; MsgBox("got excel")
		try {
			xlApp.Run(fullMacroName)
		} catch {
			MsgBox
				(
				"Error: Failed to run macro
				
				" 
				
				fullMacroName 
				
				"
	
				Possible causes:
	
				- Excel not open
				- or wrong workbookname or not open
				- or module not found
				- or macro not found 
				- or Project could not compile 
				- or wrong number/type of arguments"
				)
		}
	}

	Close(*)
	{
		delete := 0
		for i, Gui in MultiGui.Guis
		{
			if this.MyGui.Hwnd = Gui.Hwnd
			{
				MultiGui.Guis.RemoveAt(i)
				MyGui := Gui
				; MyMonitor := Gui.Monitor
				try WinGetPos(&x,&y,&w,&h,MultiGui.Guis[i-1])
				catch
					try WinGetPos(&x,&y,&w,&h,MultiGui.Guis[i])
				delete := 1
			}
		}
		if Delete = 0
			return
		WinGetPos(&ix,&iy,&w,&h,MyGui)
		loop 50
		{
			If (!Mod(A_index, 18))
			{
				WinSetTransColor("Blue " 255 - A_index * 5,MyGui)
				MyGui.Move(iX += 10, iY)
				sleep 50
			}
		}

		this.MyGui.Destroy()
		; if IsSet(x)
		; 	MultiGui.LastPOs[MyMonitor] := {x:x,y:y,w:w,h:h}
		; else
		; 	MultiGui.LastPOs[MyMonitor] := {x:0,0:0,w:0,h:0}
	}

	; Static GetTaskBarPos()
	; {
	; 	WinWait("ahk_class Shell_TrayWnd") ; incase windows starting and script load before taskbar exist then wait
	; 	WinGetPos(&x,&y,&w,&h, "ahk_class Shell_TrayWnd")
	; 	if x = 0 && y = 0 && w = A_ScreenWidth
	; 		Docked := "T"
	; 	else if x = 0 && y = 0 && h = A_ScreenHeight
	; 		Docked := "L"
	; 	else if x = 0 &&  y > 0 && w = A_ScreenWidth
	; 		Docked := "B"
	; 	else if x > 0 && y = 0 && h = A_ScreenHeight
	; 		Docked := "R"
	; 	return {x:x,y:y,w:w,h:h,Docked:Docked}
	; }

	Static CalcMonitor()
	{
		Monitors := []
		Loop  MonitorGetCount()
		{
			MonitorGetWorkArea A_Index, &L, &T, &R, &B
			Monitors.Push(Map('x',L,'y',T,'w',R,'h',B))
		}
		return Monitors
	}
	Static LastNotifyDisplay(notifyGui)
	{
		;CoordMode("Mouse","Screen")
		;MouseGetPos(&mx,&my,)
		WinGetPos(&mx,&my,,,notifyGui)
		Loop MonitorGetCount()
		{
			MonitorGet(a_index, &Left, &Top, &Right, &Bottom)
			if (Left <= mx && mx <= Right && Top <= my && my <= Bottom)
				Return MonitorGetName(a_index) ; DisplayPath[MonitorGetName(a_index)]
		}
		Return 1
	}
}


Excel_Get(WinTitle := "ahk_class XLMAIN") {	; by Sean and Jethrow, minor modification by Learning one
    hwnd := ControlGethwnd("Excel71", WinTitle)
    if !hwnd
        return
    Window := Acc_ObjectFromWindow(hwnd, -16)
    Loop
        try
            oExcel := Window.Application
        catch
            ControlSend("{esc}", "Excel71", WinTitle)
    Until !!oExcel
    return oExcel
}


Acc_ObjectFromWindow(hWnd, idObject := -4) {	; OBJID_WINDOW:=0, OBJID_CLIENT:=-4
    Acc_Init()
    capIID := 16
    bIID := Buffer(capIID)
    idObject &= 0xFFFFFFFF
    numberA := idObject == 0xFFFFFFF0 ? 0x0000000000020400 : 0x11CF3C3D618736E0
    numberB := idObject == 0xFFFFFFF0 ? 0x46000000000000C0 : 0x719B3800AA000C81
    addrPostIID := NumPut("Int64", numberA, bIID)
    addrPPIID := NumPut("Int64", numberB, addrPostIID)
    gotObject := DllCall("oleacc\AccessibleObjectFromWindow"
        , "Ptr", hWnd
        , "UInt", idObject
        , "Ptr", -capIID + addrPPIID
        , "Ptr*", &pacc := 0
    )
    if (gotObject = 0) {
        return ComObjFromPtr(pacc)
    }
}

Acc_Init() {
    static h := 0
    If Not h {
        h := DllCall("LoadLibrary", "Str", "oleacc", "Ptr")
    }
}