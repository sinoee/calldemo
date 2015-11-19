;脚本环境 Autohotkey L ANSI x32 1.1.21+
;初始化状态
#MaxMem 100
if A_IsUnicode
{
	MsgBox 4112, 环境错误, 请使用AutoHotkey L ANSI 32位版本测试或编译！
	ExitApp
}
bt := {"call":1}
;GotMessages := []
islogin := 0
callstatus = 空闲

Menu, Tray, NoStandard
Menu, Tray, add, 拒接/挂断, hangup
Menu, Tray, add, 保持, hold
Menu, Tray, add, 三方, conference
Menu, Tray, add, 转接, transfer
Menu, Tray, add, 转分机, Senddtmf
Menu, Tray, add, 小休, rest
Menu, Tray, add, 空闲, idle
Menu, Tray, add, 监听, monitor
Menu, Tray, add, 静音, mute
Menu, Tray, add, 注销, logout
Menu, Tray, add, 设置, settings
Menu, Tray, Add, 退出, loginGuiClose

Gui, Monitor:Show, , AHK_Monitor
Gui, Monitor:+hwndMonitorHwnd
Gui, Monitor:+LastFound
OnMessage(0x4A,"Receive_WM_COPYDATA")
Gui, Monitor:Hide

Process, Exist, WinvoiceCC.exe ;运行控件
if !ErrorLevel
	Run, C:\WinvoiceCC\WinvoiceCC.exe

Process, Wait, WinvoiceCC.exe

Run, getmessage.exe

;goto, showmainui
IfExist, keep.ini
{
	IniRead, user, keep.ini, login, user
	IniRead, pass, keep.ini, login, pass
}
Gui, login:Add, Picture, x0 y0, user.png
Gui, login:Add, Edit, x20 y0 w200 h20 vuser, % user
Gui, login:Add, Picture, x0 y20, key.png
Gui, login:Add, Edit, x20 y20 w200 h20 Password vpass, % pass
Gui, login:Add, Checkbox, x0 y40 w80 h20 vkeeppass, 记住密码
Gui, login:Add, Button, x0 y60 w220 h20 glogin Default, 安全登录
Gui, login:show, , 用户登录
return

login:
	Gui, login:Submit, NoHide
	;注册成功消息
	;{"header":105,"code":0,"value":{"voice_agentid":"88893","voice_pwd":"","voice_ip":"122.144.133.56","voice_port":"5060","prefix":"501","company_id":"21","company_code":"000002"}}
	;注册失败消息
	;{"header":106,"code":1002,"value":""}
	if !keeppas
		FileDelete, keep.ini
	reg_str := "http://localhost:8889/login?userid=" user "&password=" pass "&flag=1"
	reg_back := json_toobj(URLDownloadToVar(reg_str))
	/*
	while (!(Message := GotMessages.Pop())) ;等待消息
	Sleep, 100
	if (Message.header = 105)
	{
	Gui, login:Destroy
	goto, showmainui
	}
	else
	{
	MsgBox, 4112, 错误, 用户名或密码错误！
	}
	*/
	while (!islogin) ;等待消息
		Sleep, 100
	if (result.header = 105)
	{
		Gui, login:Destroy
		goto, showmainui
	}
	else
	{
		MsgBox, 4112, 错误, 用户名或密码错误！
	}
return


showmainui:
	Gui, main:Font, s13 , Arial
	Gui, main:Add, edit, x2 y2 w130 h27 -Multi vphone,
	Gui, main:Add, Picture, x130 y2 w30 h28 gcall vcall, % bt.call ? "call.png" : "hangup.png"
	Gui, main: -Caption +ToolWindow +AlwaysOnTop +LastFound +Hwndmainhwnd
	Gui, main:Default
	Gui, main:Show, % "x" A_ScreenWidth-240 " y80 w162 h32",
	OnMessage(0x204, "WM_RButtonDOWN") ;监听右键

	;右键菜单

	Menu, lb_Menu, add, 拒接/挂断, hangup
	Menu, lb_Menu, add, 保持, hold
	Menu, lb_Menu, add, 三方, conference
	Menu, lb_Menu, add, 转接, transfer
	Menu, lb_Menu, add, 转分机, Senddtmf
	Menu, lb_Menu, add, 小休, rest
	Menu, lb_Menu, add, 空闲, idle
	Menu, lb_Menu, add, 监听, monitor
	Menu, lb_Menu, add, 静音, mute
	Menu, lb_Menu, add, 注销, logout
	Menu, lb_Menu, add, 设置, settings
return

WM_RButtonDOWN()
{
	if A_GuiControl <> "phone"
		Menu, lb_Menu, show
}

test:
	ToolTip % A_ThisMenuItem
return

hangup:
	hangup_back =
	while(!hangup_back) ;挂不断 再试几次 汗
		hangup_back := json_toobj(URLDownloadToVar("http://localhost:8889/hangup"))
idle:
	idle_back := json_toobj(URLDownloadToVar("http://localhost:8889/idle"))
	callstatus = 空闲
	bt.call := 1
	GuiControl, main:, call, call.png
return


call:
	GuiControlGet, phone
	MsgBox % bt.call
	if bt.call
	{
		if callstatus = 空闲
		{
			call_back := json_toobj(URLDownloadToVar("http://localhost:8889/outbound?dst=" phone))
			callstatus = 呼出
		}
		else if callstatus = 来电
		{
			call_back := json_toobj(URLDownloadToVar("http://localhost:8889/answer"))
		}
	}
	else
	{
		hangup_back =
		while(!hangup_back) ;挂不断 再试几次 汗
			hangup_back := json_toobj(URLDownloadToVar("http://localhost:8889/hangup"))
		idle_back := json_toobj(URLDownloadToVar("http://localhost:8889/idle"))
		callstatus = 空闲

	}
	bt.call := !bt.call
	GuiControl, main:, call,  % bt.call ? "call.png" : "hangup.png"
return

phone:
	GuiControlGet, phone
	if RegExMatch(phone,"[\s\-]+")
	{
		GuiControl, main:, phone, % RegExReplace(phone,"[\s\-]+","")
		ControlSend, Edit1, {End}, ahk_id %mainhwnd%
	}
return

hold:


conference:


transfer:


senddtmf:


rest:

monitor:

mute:

settings:


return

popui:
	if !popui
	{
		WinGetPos, , , , tray_h, ahk_class Shell_TrayWnd
		Gui, popui:Font, s16 Bold cBlack, Arial
		Gui, popui:Add, Text, xm y0 w200 h30 vpopphone,
		Gui, popui:Add, Button, x2 y30 w98 h20, 接听
		Gui, popui:Add, Button, x102 y30 w98 h20, 拒接
		Gui, popui: +ToolWindow +AlwaysOnTop +hwndpopuihwnd
		Gui, popui:Show, x-300 y-300 w200 NoActivate
		WinGetPos, , , pop_w, pop_h, ahk_id %popuihwnd%
		Gui, popui:Show, % "x" A_ScreenWidth-pop_w " y" A_ScreenHeight-tray_h-pop_h, 来电
		popui = 1
	}
	else
		Gui, popui:Show
return

popuiGuiclose:
	Gui, popui:Hide
return

logout:
mainGuiClose:
	logout_back := json_toobj(URLDownloadToVar("http://localhost:8889/logout"))
loginGuiClose:
	Process, Close, getmessage.exe
	Process, Close, WinvoiceCC.exe
	ExitApp


#If WinActive("ahk_id " mainhwnd)

Enter::
	goto, call
return

~LButton:: ;拖拽触发
	KeyWait, LButton, T0.2
	if !ErrorLevel
		Send, {LButton}
	else
	{
		CoordMode, Mouse
		MouseGetPos, MouseStartX, MouseStartY
		WinGetPos, OriginalPosX, OriginalPosY,,, ahk_id %mainhwnd%
		WinGet, WinState, MinMax, ahk_id %mainhwnd%
		if WinState = 0
			SetTimer, WatchMouse, 10
	}
return

;音量调节
WheelUp::

WheelDown::

return


;按键音
~0::
~1::
~2::
~3::
~4::
~5::
~6::
~7::
~8::
~9::
	SoundPlay, % RegExReplace(A_ThisHotkey,"~","") ".wav"
return

#If


WatchMouse: ;拖拽定时器任务
	GetKeyState, LButtonState, LButton, P
	if LButtonState = U
	{
		SetTimer, WatchMouse, off
		return
	}
	GetKeyState, EscapeState, Escape, P
	if EscapeState = D
	{
		SetTimer, WatchMouse, off
		WinMove, ahk_id %mainhwnd%,, %OriginalPosX%, %OriginalPosY%
		return
	}
	CoordMode, Mouse
	MouseGetPos, MouseX, MouseY
	WinGetPos, WinX, WinY,,, ahk_id %mainhwnd%
	SetWinDelay, -1
	WinMove, ahk_id %mainhwnd%,, WinX + MouseX - MouseStartX, WinY + MouseY - MouseStartY
	MouseStartX := MouseX
	MouseStartY := MouseY
return


show_obj(obj,Menu_name:=""){ ;调试输出数组
static id
if Menu_name =
{
	main = 1
	id++
	Menu_name := id
}
Menu, % Menu_name, add,
Menu, % Menu_name, DeleteAll
for k,v in obj
{
	if (IsObject(v))
	{
		id++
		subMenu_name := id
		Menu, % subMenu_name, add,
		Menu, % subMenu_name, DeleteAll
		Menu, % Menu_name, add, % k ? "【" k "】[obj]" : "", :%subMenu_name%
		show_obj(v,subMenu_name)
	}
	else
	{
		Menu, % Menu_name, add, % k ? "【" k "】" v: "", MenuHandler
	}
}
if main = 1
	Menu,% Menu_name, show

MenuHandler:
return
}

Deal(Code,Msg)
{
	global
	result := json_toobj(Msg)
	;GotMessages.Push(result)
	if (result.header = 105) ;注册成功
		islogin := 1
	else if (result.header = 106) ;注册失败
		islogin := -1
	else if (result.header = 101) ;来电
	{
		callstatus = 来电
		gosub, popui
		GuiControl, popui:, popphone, % result.value
	}
	else if (result.header = 102) ;振铃
	{
		callstatus = 振铃
	}
	else if (result.header = 103) ;接通
	{
		callstatus = 接通
	}
	else if (result.header = 104) ;挂机
	{
		callstatus = 空闲
		bt.call := 1
		GuiControl, main:, call, call.png
	}
	else if (result.header = 107) ;网络信息情况
	{

	}

	return 1
}

Receive_WM_COPYDATA(wParam, lParam)
{
	StringAddress := NumGet(lParam + 2*A_PtrSize) ;获取文本指针
	length := NumGet(lParam + A_PtrSize)
	CopyOfData := StrGet(StringAddress,length) ;指针到文本
	return Deal(wParam,CopyOfData) ;消息号，信息
}


URLDownloadToVar(url, Encoding = "",Method="GET",postData=""){ ;网址，编码,请求方式，post数据
hObject:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
if Method = GET
{
	Try
	{
		hObject.Open("GET",url)
		hObject.Send()
	}
	catch e
		return -1
}
else if Method = POST
{
	Try
	{
		hObject.Open("POST",url,False)
		hObject.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
		hObject.Send(postData)
	}
	catch e
		return -1
}

if (Encoding && hObject.ResponseBody)
{
	oADO := ComObjCreate("adodb.stream")
	oADO.Type := 1
	oADO.Mode := 3
	oADO.Open()
	oADO.Write(hObject.ResponseBody)
	oADO.Position := 0
	oADO.Type := 2
	oADO.Charset := Encoding
	return oADO.ReadText(), oADO.Close()
}
return hObject.ResponseText
}



Ansi2UTF8(sString) ;GBK转UTF-8
{
	Ansi2Unicode(sString, wString, 0)
	Unicode2Ansi(wString, zString, 65001)
	return zString
}
UTF82Ansi(zString) ;UTF-8转GBK
{
	Ansi2Unicode(zString, wString, 65001)
	Unicode2Ansi(wString, sString, 0)
	return sString
}
Ansi2Unicode(ByRef sString, ByRef wString, CP = 0)
{
	nSize := DllCall("MultiByteToWideChar"
		, "Uint", CP
		, "Uint", 0
		, "Uint", &sString
		, "int", -1
		, "Uint", 0
		, "int", 0)
	VarSetCapacity(wString, nSize * 2)
	DllCall("MultiByteToWideChar"
		, "Uint", CP
		, "Uint", 0
		, "Uint", &sString
		, "int", -1
		, "Uint", &wString
		, "int", nSize)
}
Unicode2Ansi(ByRef wString, ByRef sString, CP = 0)
{
	nSize := DllCall("WideCharToMultiByte"
		, "Uint", CP
		, "Uint", 0
		, "Uint", &wString
		, "int", -1
		, "Uint", 0
		, "int", 0
		, "Uint", 0
		, "Uint", 0)
	VarSetCapacity(sString, nSize)
	DllCall("WideCharToMultiByte"
		, "Uint", CP
		, "Uint", 0
		, "Uint", &wString
		, "int", -1
		, "str", sString
		, "int", nSize
		, "Uint", 0
		, "Uint", 0)
}

urlencode(string){
string := Ansi2UTF8(string)
StringLen, len, string
Loop % len
{
	SetFormat, IntegerFast, hex
	StringMid, out, string, %A_Index%, 1
	hex := Asc(out)
	hex2 := hex
	StringReplace, hex, hex, 0x, , All
	SetFormat, IntegerFast, d
	hex2 := hex2
	if (hex2==33 || (hex2>=39 && hex2 <=42) || hex2==45 || hex2 ==46 || (hex2>=48 && hex2<=57) || (hex2>=65 && hex2<=90) || hex2==95 || (hex2>=97 && hex2<=122) || hex2==126)
		content .= out
	else
		content .= "`%" hex
}
return content
}

json_fromobj( obj ) {

if IsObject( obj )
{
	isarray := 0 ; an empty object could be an array... but it ain't, says I
	for key in obj
		if ( key != ++isarray )
		{
			isarray := 0
			break
		}

	for key, val in obj
		str .= ( A_Index = 1 ? "" : "," ) ( isarray ? "" : json_fromObj( key ) ":" ) json_fromObj( val )

	return isarray ? "[" str "]" : "{" str "}"
}
else if obj IS NUMBER
	return obj
;	else if obj IN null,true,false ; AutoHotkey does not natively distinguish these
;		return obj

; Encode control characters, starting with backslash.
StringReplace, obj, obj, \, \\, A
StringReplace, obj, obj, % Chr(08), \b, A
StringReplace, obj, obj, % A_Tab, \t, A
StringReplace, obj, obj, `n, \n, A
StringReplace, obj, obj, % Chr(12), \f, A
StringReplace, obj, obj, `r, \r, A
StringReplace, obj, obj, ", \", A
StringReplace, obj, obj, /, \/, A
while RegExMatch( obj, "[^\x20-\x7e]", key )
{
	str := Asc( key )
	val := "\u" . Chr( ( ( str >> 12 ) & 15 ) + ( ( ( str >> 12 ) & 15 ) < 10 ? 48 : 55 ) )
	. Chr( ( ( str >> 8 ) & 15 ) + ( ( ( str >> 8 ) & 15 ) < 10 ? 48 : 55 ) )
	. Chr( ( ( str >> 4 ) & 15 ) + ( ( ( str >> 4 ) & 15 ) < 10 ? 48 : 55 ) )
	. Chr( ( str & 15 ) + ( ( str & 15 ) < 10 ? 48 : 55 ) )
	StringReplace, obj, obj, % key, % val, A
}
return """" obj """"
}

json_toobj(str){

quot := """" ; firmcoded specifically for Readability. Hardcode for (minor) performance gain
ws := "`t`n`r " Chr(160) ; whiteSpace plus NBSP. This gets trimmed from the markup
obj := {} ; dummy object
objs := [] ; stack
keys := [] ; stack
isarrays := [] ; stack
literals := [] ; queue
y := nest := 0

; First pass swaps out literal strings so we can parse the markup easily
StringGetPos, z, str, %quot% ; initial seek
while !ErrorLevel
{
	; Look for the non-literal quote that ends this string. Encode literal backslashes as '\u005C' because the
	; '\u..' entities are decoded last and that prevents literal backslashes from borking normal characters
	StringGetPos, x, str, %quot%,, % z + 1
	while !ErrorLevel
	{
		StringMid, key, str, z + 2, x - z - 1
		StringReplace, key, key, \\, \u005C, A
		if SubStr( key, 0 ) != "\"
			break
		StringGetPos, x, str, %quot%,, % x + 1
	}
	;	StringReplace, str, str, %quot%%t%%quot%, %quot% ; this might corrupt the string
	str := ( z ? SubStr( str, 1, z ) : "" ) quot SubStr( str, x + 2 ) ; this won't

	; Decode entities
	StringReplace, key, key, \%quot%, %quot%, A
	StringReplace, key, key, \b, % Chr(08), A
	StringReplace, key, key, \t, % A_Tab, A
	StringReplace, key, key, \n, `n, A
	StringReplace, key, key, \f, % Chr(12), A
	StringReplace, key, key, \r, `r, A
	StringReplace, key, key, \/, /, A
	while y := InStr( key, "\u", 0, y + 1 )
	if ( A_IsUnicode || Abs( "0x" SubStr( key, y + 2, 4 ) ) < 0x100 )
		key := ( y = 1 ? "" : SubStr( key, 1, y - 1 ) ) Chr( "0x" SubStr( key, y + 2, 4 ) ) SubStr( key, y + 6 )

	literals.Insert(key)

	StringGetPos, z, str, %quot%,, % z + 1 ; seek
}

; Second pass parses the markup and builds the object iteratively, swapping placeholders as they are encountered
key := isarray := 1

; The outer loop splits the blob into paths at markers where nest level decreases
Loop Parse, str, % "]}"
{
	StringReplace, str, A_LoopField, [, [], A ; mark any array open-brackets

	; This inner loop splits the path into segments at markers that signal nest level increases
	Loop Parse, str, % "[{"
	{
		; The first segment might contain members that belong to the previous object
		; Otherwise, push the previous object and key to their stacks and start a new object
		if ( A_Index != 1 )
		{
			objs.Insert( obj )
			isarrays.Insert( isarray )
			keys.Insert( key )
			obj := {}
			isarray := key := Asc( A_LoopField ) = 93
		}

		; arrrrays are made by pirates and they have index keys
		if ( isarray )
		{
			Loop Parse, A_LoopField, `,, % ws "]"
				if ( A_LoopField != "" )
					obj[key++] := A_LoopField = quot ? literals.Remove(1) : A_LoopField
		}
		; otherwise, parse the segment as key/value pairs
		else
		{
			Loop Parse, A_LoopField, `,
				Loop Parse, A_LoopField, :, % ws
					if ( A_Index = 1 )
						key := A_LoopField = quot ? literals.Remove(1) : A_LoopField
					else if ( A_Index = 2 && A_LoopField != "" )
						obj[key] := A_LoopField = quot ? literals.Remove(1) : A_LoopField
		}
		nest += A_Index > 1
	} ; Loop Parse, str, % "[{"

	if !--nest
		break

	; Insert the newly closed object into the one on top of the stack, then pop the stack
	pbj := obj
	obj := objs.Remove()
	obj[key := keys.Remove()] := pbj
	if ( isarray := isarrays.Remove() )
		key++

} ; Loop Parse, str, % "]}"

return obj
}

