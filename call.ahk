;脚本环境 Autohotkey L ANSI x32 1.1.21+
;初始化状态
#maxmem 100

bt := {"call":1,"keep":0,"xfer":0}
GotMessages := []
Menu, Tray, NoStandard
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

;去掉注释隐藏运行
Run, getmessage.exe , , Hide
;Run, getmessage.exe ;, , Hide

;goto, showmainui
IfExist, keep.ini
{
	IniRead, user, keep.ini, login, user
	IniRead, pass, keep.ini, login, pass
}
Gui, login:font, s10, Verdana ;设置字体
Gui, login:Add, Picture, x0 y5, user.png
Gui, login:Add, Edit, x30 y5 w200 h20 vuser, % user
Gui, login:Add, Picture, x0 y30, key.png
Gui, login:Add, Edit, x30 y30 w200 h20 Password vpass, % pass
Gui, login:Add, Checkbox, x0 y50 w100 h30 vkeeppass, 记住密码
Gui, login:Add, Checkbox, x100 y50 w100 h30 vqutologin, 自动登录
Gui, login:Add, Button, x0 y83 w230 h20 glogin Default, 安全登录
Gui, login:show, , 用户登录
return

login:
Gui, login:Submit, NoHide
;http://localhost:8889/login?userid=abc&password=abc123&flag=1
;注册成功消息
;{"header":105,"code":0,"value":{"voice_agentid":"88893","voice_pwd":"","voice_ip":"122.144.133.56","voice_port":"5060","prefix":"501","company_id":"21","company_code":"000002"}}
;注册失败消息
;{"header":106,"code":1002,"value":""}
if !keeppas
	FileDelete, keep.ini
reg_str := "http://localhost:8889/login?userid=" user "&password=" pass "&flag=1"
reg_back := json_toobj(URLDownloadToVar(reg_str))
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
return


showmainui:
Gui, main:font, s10, Verdana ;设置字体
;Gui, main:Flash  ;来电时编辑框变换颜色
Gui, main:Add, edit, x2 y2 w148 h26 -Multi vphone,
Gui, main:Add, Picture, x148 y1 w30 h27 gcall vcall, % bt.call ? "call.png" : "hangup.png"
Gui, main: -Caption +ToolWindow +AlwaysOnTop +LastFound +Hwndmainhwnd +Owner
Gui, main:Default
Gui, main:Show, % "x" A_ScreenWidth-240 " y80 w180 h32",
;Gui, main:Show, w180 h32,
return

call:
GuiControlGet, phone
;MsgBox % bt.call
if bt.call
{
	call_back := json_toobj(URLDownloadToVar("http://localhost:8889/outbound?dst=" phone "&callback="))

}
else
{
	hangup_back =
	while(!hangup_back) ;挂不断 再试几次 汗
		hangup_back := json_toobj(URLDownloadToVar("http://localhost:8889/hangup"))

}
bt.call := !bt.call
GuiControl, main:, call,  % bt.call ? "call.png" : "hangup.png"
return


loginGuiClose:
mainGuiClose:
Process, Close, getmessage.exe
Process, Close, WinvoiceCC.exe
ExitApp


#if winactive("ahk_id " mainhwnd) ;拖拽触发
~LButton::
Sleep, 200
if GetKeyState("LButton","P")
{
	CoordMode, Mouse
	MouseGetPos, MouseStartX, MouseStartY
	WinGetPos, OriginalPosX, OriginalPosY,,, ahk_id %mainhwnd%
	WinGet, WinState, MinMax, ahk_id %mainhwnd%
	if WinState = 0
		SetTimer, WatchMouse, 10
}
return
#if

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


show_obj(obj,menu_name:=""){ ;调试输出数组
static id
if menu_name =
    {
    main = 1
    id++
    menu_name := id
    }
Menu, % menu_name, add,
Menu, % menu_name, DeleteAll
for k,v in obj
{
if (IsObject(v))
	{
    id++
    submenu_name := id
    Menu, % submenu_name, add,
    Menu, % submenu_name, DeleteAll
	Menu, % menu_name, add, % k ? "【" k "】[obj]" : "", :%submenu_name%
    show_obj(v,submenu_name)
	}
Else
	{
	Menu, % menu_name, add, % k ? "【" k "】" v: "", MenuHandler
	}
}
if main = 1
    menu,% menu_name, show

MenuHandler:
return
}

Deal(Code,Msg)
{
global GotMessages
result := json_toobj(Msg)
GotMessages.Push(result)
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
   Return zString
}
UTF82Ansi(zString) ;UTF-8转GBK
{
   Ansi2Unicode(zString, wString, 65001)
   Unicode2Ansi(wString, sString, 0)
   Return sString
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
If (hex2==33 || (hex2>=39 && hex2 <=42) || hex2==45 || hex2 ==46 || (hex2>=48 && hex2<=57) || (hex2>=65 && hex2<=90) || hex2==95 || (hex2>=97 && hex2<=122) || hex2==126)
	content .= out
Else
	content .= "`%" hex
}
Return content
}

json_fromobj( obj ) {

	If IsObject( obj )
	{
		isarray := 0 ; an empty object could be an array... but it ain't, says I
		for key in obj
			if ( key != ++isarray )
			{
				isarray := 0
				Break
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
	While RegexMatch( obj, "[^\x20-\x7e]", key )
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

	quot := """" ; firmcoded specifically for readability. Hardcode for (minor) performance gain
	ws := "`t`n`r " Chr(160) ; whitespace plus NBSP. This gets trimmed from the markup
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
			If SubStr( key, 0 ) != "\"
				Break
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

		literals.insert(key)

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
				objs.insert( obj )
				isarrays.insert( isarray )
				keys.insert( key )
				obj := {}
				isarray := key := Asc( A_LoopField ) = 93
			}

			; arrrrays are made by pirates and they have index keys
			if ( isarray )
			{
				Loop Parse, A_LoopField, `,, % ws "]"
					if ( A_LoopField != "" )
						obj[key++] := A_LoopField = quot ? literals.remove(1) : A_LoopField
			}
			; otherwise, parse the segment as key/value pairs
			else
			{
				Loop Parse, A_LoopField, `,
					Loop Parse, A_LoopField, :, % ws
						if ( A_Index = 1 )
							key := A_LoopField = quot ? literals.remove(1) : A_LoopField
						else if ( A_Index = 2 && A_LoopField != "" )
							obj[key] := A_LoopField = quot ? literals.remove(1) : A_LoopField
			}
			nest += A_Index > 1
		} ; Loop Parse, str, % "[{"

		If !--nest
			Break

		; Insert the newly closed object into the one on top of the stack, then pop the stack
		pbj := obj
		obj := objs.remove()
		obj[key := keys.remove()] := pbj
		If ( isarray := isarrays.remove() )
			key++

	} ; Loop Parse, str, % "]}"

	Return obj
}


HashFromAddr(pData, len, algid, key=0)
{
  hProv := size := hHash := hash := ""
  ptr := (A_PtrSize) ? "ptr" : "uint"
  aw := (A_IsUnicode) ? "W" : "A"
  if (DllCall("advapi32\CryptAcquireContext" aw, ptr "*", hProv, ptr, 0, ptr, 0, "uint", 1, "uint", 0xF0000000))
  {
    if (DllCall("advapi32\CryptCreateHash", ptr, hProv, "uint", algid, "uint", key, "uint", 0, ptr "*", hHash))
    {
      if (DllCall("advapi32\CryptHashData", ptr, hHash, ptr, pData, "uint", len, "uint", 0))
      {
        if (DllCall("advapi32\CryptGetHashParam", ptr, hHash, "uint", 2, ptr, 0, "uint*", size, "uint", 0))
        {
          VarSetCapacity(bhash, size, 0)
          DllCall("advapi32\CryptGetHashParam", ptr, hHash, "uint", 2, ptr, &bhash, "uint*", size, "uint", 0)
        }
      }
      DllCall("advapi32\CryptDestroyHash", ptr, hHash)
    }
    DllCall("advapi32\CryptReleaseContext", ptr, hProv, "uint", 0)
  }
  int := A_FormatInteger
  SetFormat, Integer, h
  Loop, % size
  {
    v := substr(NumGet(bhash, A_Index-1, "uchar") "", 3)
    while (strlen(v)<2)
      v := "0" v
    hash .= v
  }
  SetFormat, Integer, % int
  return hash
}


HashFromString(string, algid, key=0)
{
  len := strlen(string)
  if (A_IsUnicode)
  {
    ;VarSetCapacity(data, len)
    ;StrPut(string, &data, len, "cp0")
    return HashFromAddr(&data, len, algid, key)
  }
  data := string
  return HashFromAddr(&data, len, algid, key)
}

MD5(string,b16:=false) ;b16 是否16位
{
  ;0x8003/*_CALG_MD5*/
  ;0x8001/*_CALG_MD2*/
  ;0x8002/*_CALG_MD4*/
  ;0x8004/*_CALG_SHA1*/
  return b16 ? SubStr(HashFromString(string, 0x8003),9,16) : HashFromString(string, 0x8003)
}
