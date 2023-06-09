;# = Windows Key
;^ = Control
;! = Alt
;+ = Shift

;:?*:``:: way to do double tap hotkey

#SingleInstance force ;allows reloading the script w/o popup window
#Persistent

;Start of Auto-Execute section

;win group to enter number with wheeldown

	GroupAdd, InputNumber, Image Position
	GroupAdd, InputNumber, Score: Trade Execution
	GroupAdd, InputNumber, Score: Trade Performance

;provides a reminder to save MT4 statement 3 days before end of month

	LDM := LDOM()			;calls function to acquire the last date of this month
	FormatTime wkday, %LDM%, wday	;acquires the weekday of the last day of this month
	FormatTime cMnth, %LDM%, MMMM	;acquires the long name of the month
	tDate := SubStr(A_Now, 1, 8)	;todays date
	rDays := LDM			;copy last day of month to find difference of today
	rDays -= tDate, d		;acquire remaining days in the month
	Gosub, L3DM			;acquires the last 3 days of month excluding Saturdays

	loop, 3				;checks today's date against the last 3 days of month
	 {				;provides a msgbox reminder if within the last 3 days
	  i = %A_Index%
	  if % tDate = rDate%i%
	   {
	    msgbox, 262144, Reminder: Save MT4 Statement, %Reminder%
	    break
	   }
	 }

;closes update link windows in journal

	SetTimer, CloseUpdateWin

	CloseUpdateWin:

	IfWinActive, Microsoft Excel ahk_class #32770, Remote data not accessible
	 send, {right 3}{enter}

	return

;End of Auto-Execute section

;Allows escape from this script

	End::ExitApp	;exit script

;Name of excel journal for this entire script

	Title:
	JournalTitle := "Caissia - Investment.xlsm"
	return

;Exit scripts

	Release:
	suspend, off
	BlockInput, off
	BlockInput, MouseMoveOff
	return

#IfWinActive ahk_class Shell_TrayWnd
;Open  MetaTrader 4 and Trade Journal Excel, ScreenHunter
;Close MetaTrader 4 and Trade Journal Excel, ScreenHunter

	mbutton::
	^+m::

	keywait, mbutton
	keywait, m

	MouseGetPos, xm, ym

	;must be on last icon to move it next to excel journal
	if (xm > 1474) and (xm < 1516) and (ym > 0) and (ym < 32)
	 {
	  blockinput, mousemove
	  click down
	  mousemove 400,20
	  click up
	  blockinput, mousemoveoff
	  return
	 }

	;must be on taskbar's arrow next to "Terminal" to close apps
	if (xm > 1650) and (xm < 1664) and (ym > 0) and (ym < 32)
	 {
	  Process, close, ScreenHunter.exe	
	  WinClose, ahk_class MetaQuotes::MetaTrader::4.00
	  WinActivate, ahk_class XLMAIN
	  WinClose, ahk_class XLMAIN
	  return
	 }

	;must be on taskbar's "Terminal" to toggle or open apps
	if (xm < 1584) or (xm > 1644) or (ym < 0) or (ym > 32)
	 {
	  return
	 }

	TradeToggle:

	Gosub, Title

	Process, exist, ScreenHunter.exe
	 if ErrorLevel = 0
	  {
	   run, C:\Users\image\Documents\trade\ScreenHunter
	   ScreenHunter := " ScreenHunter "
	   WinWaitActive, %ScreenHunter%, , 1
	   WinClose, %ScreenHunter%
	  }

	IfWinNotExist, ahk_class MetaQuotes::MetaTrader::4.00
	 {
	  run, C:\Users\image\Documents\trade\FXCM,, max
	  sleep, 4000
	  WinActivate, ahk_class XLMAIN
	  sleep, 1000
	  send, ^+r	;macro to refresh DDE link to MT4
	  IfWinExist %JournalTitle%
	    return
	 }

	IfWinNotExist %JournalTitle%
	 {
	  run, C:\Users\image\Documents\trade\Investment,, max
	  sleep, 2000
	  WinActivate, ahk_class XLMAIN
	  WinWaitClose, Ultimate Space
	  WinActivate, ahk_class MetaQuotes::MetaTrader::4.00
	  return
	 }

	TradeToggle := !TradeToggle

	if TradeToggle
	 {
	  WinActivate, ahk_class MetaQuotes::MetaTrader::4.00
	  return
	 }

	if !TradeToggle
	 {
	  WinActivate, %JournalTitle%
	  return
	 }

	return

#IfWinActive


#IfWinActive, ahk_class XLMAIN
;Open/toggle to MT4

	mbutton::
	^+m::
	
	keywait, mbutton
	keywait, m
	MouseGetPos, xm, ym

	;if mouse is over taskbar area then do nothing
	if (ym > 1050)
	 return
	else
	 Gosub, TradeToggle
	return


;Open Vision message

	^mbutton::
	send, ^+U
	return

#IfWinActive


#IfWinActive ahk_class MetaQuotes::MetaTrader::4.00

;Hotkeys to send to MT4 to activate MQL4 script

	Tab & e::		;next timeframe for all windows
	umDelay := 1
	suspend, on
	send, !1
	suspend, off
	return

	Tab & s::		;input timeframe for all windows
	send, !2
	return

	Tab & q::		;previous timeframe for all windows
	umDelay := 1
	suspend, on
	send, !3
	suspend, off
	return

	Tab & d::		;next timeframe
	send, !4
	return

	Tab & w::		;resets & sets up subwindows to scroll from left to right
	umDelay := 1
	suspend, on
	send, !5
	suspend, off
	return

	Tab & a::		;previous timeframe
	send, !6
	return

	Tab & z::		;close all sub-windows except for focus window
	send, !u
	return

	Capslock & e::		;next template for all windows
	umDelay := 1
	suspend, on		;required since focus is moved off of MT4
	send, !7
	sleep, 200
	suspend, off
	return

	Capslock & w::		;input template for all windows
	send, !8
	sleep, 200
	WinWaitClose, Template_All ahk_class #32770
	sleep, 200
	return

	Capslock & q::		;previous template for all windows
	umDelay := 1
	suspend, on
	send, !9
	sleep, 200
	suspend, off
	return

	Capslock & d::		;next template
	umDelay := 1
	suspend, on		;required since random hks are activated as well
	sendinput, !q
	suspend, off
	sleep, 200
	return

	Capslock & s::		;next chart window
	send, !w
	return

	Capslock & a::		;previous template
	umDelay := 1
	suspend, on		;required since random hks are activated as well
	sendinput, !e
	suspend, off
	sleep, 200
	return

	Capslock & g::		;next profile
	sendinput, ^{F5}
	return

	Capslock & f::		;previous profile
	sendinput, +{F5}
	return

	Capslock & c::		;clone current window
	umDelay := 1
	suspend, on		;required since random hks are activated as well
	send, !y
	suspend, off
	return

	Capslock & z::		;close current chart window
	umDelay := 1
	suspend, on
	sendinput, {Ctrl down}{f4}
	sendinput, {Ctrl up}
	suspend, off
	return

	shift & d::		;next currency pair of subset of currency pairs
	umDelay := 1
	suspend, on		;required since random hks are activated as well
	sendinput, !t
	suspend, off
	return
	
	shift & s::		;open all the symbols in subset of currency pairs (20)
	sendinput, !l
	return

	shift & a::		;previous currency pair of subset of currency pairs
	umDelay := 1
	suspend, on		;required since random hks are activated as well
	sendinput, !c
	suspend, off
	return

	shift & c::		;next currency pair
	umDelay := 1
	suspend, on		;required since random hks are activated as well
	sendinput, !i
	suspend, off
	return
	
	shift & x::		;input currency pair
	sendinput, !o
	return

	shift & z::		;previous currency pair
	umDelay := 1
	suspend, on		;required since random hks are activated as well
	sendinput, !p
	suspend, off
	return

	Control::		;toggles between templates: QuickSilver & Wave
	suspend, on		;required since random hks are activated as well
	sendinput, ^0
	suspend, off
	return

	Alt::			;toggles between templates: QuickSilver & Wave
	suspend, on		;required since random hks are activated as well
	sendinput, !D
	suspend, off
	return

;Open calculator

	:?*:cc::
	keywait, c
	run calc.exe
	return

;Fullscreen toggle

	:?*:ff::
	keywait, f
	send, {F11}
	return

;View all hotkeys

	:?*:ll::
	keywait, l
	ListHotkeys
	return

;Open new order window

	:?*:nn::
	keywait, n
	;suspend, on
	;sleep, 1000
	send, {F9}
	;suspend, off
	return

;Setup scrolling sub-windows left to right

	:?*:ww::
	MouseGetPos, xs, ys
	BlockInput, MouseMove
	suspend, on
	send, !r
	sleep, 800
	suspend, off
	BlockInput, MouseMoveOff
	MouseMove, xs, ys
	return

;Display templates and their corresponding numbers

	:?*:tt::
	keywait, t
	msgbox,	262144, Assistant, Template/Setups:		#
					.`n`n     QuickSilver		1
					.`n     Propulsion		2
					.`n     Bullet			3
					.`n     Triple			4
					.`n     Spring MA		5
					.`n     NTP Scale		6
					.`n     Candle Pattern		7
					.`n     PTP Setup		8
					.`n     Pivot Points (d)		9
					.`n     Pivot Points (w)		10
					.`n     Pivot Points (m)		11
					.`n     Convergence (a)	12
					.`n     Convergence (b)	13
					.`n     Alligator Volume	14
					.`n     Wave			15 .
	return

;Display subset of all currency pairs

	:?*:dd::
	keywait, d
	msgbox, 262144, Subset of Currency Pairs, Alphabetical Order: `n
 .........................................`n 1.	AUDJPY
				,`n 2.	AUDUSD
				,`n 3.	CADJPY
				,`n 4.	CHFJPY
				,`n 5.	EURAUD
				,`n 6.	EURCAD
				,`n 7.	EURGBP
				,`n 8.	EURJPY
				,`n 9.	EURNZD
				,`n 10.	EURUSD
				,`n 11.	GBPAUD
				,`n 12.	GBPCAD
				,`n 13.	GBPCHF
				,`n 14.	GBPJPY
				,`n 15.	GBPUSD
				,`n 16.	NZDJPY
				,`n 17.	NZDUSD
				,`n 18.	USDCAD
				,`n 19.	USDCHF
				,`n 20.	USDJPY,
	return

;Display all currency pairs and their #keys

	:?*:ee::
	keywait, e
	toggleList:=!toggleList	;toggles between alphabetical and numberal order

	if toggleList
		{
			msgbox, 262144, Currency Pairs, Alphabetial Order: `n
		       ..................`n EUR/AUD	-  7
					,`n GBP/AUD	-  3
					,`n AUD/CAD	-  22
					,`n EUR/CAD	-  9
					,`n GBP/CAD	-  4
					,`n USD/CAD	-  13
					,`n AUD/CHF	-  21
					,`n CAD/CHF	-  20
					,`n EUR/CHF	-  24
					,`n GBP/CHF	-  8
					,`n USD/CHF	-  17
					,`n EUR/GBP	-  23
					,`n AUD/JPY		-  10
					,`n CAD/JPY		-  12
					,`n CHF/JPY		-  16
					,`n EUR/JPY		-  6
					,`n GBP/JPY		-  2
					,`n NZD/JPY		-  14
					,`n USD/JPY		-  19
					,`n EUR/NZD	-  1
					,`n AUD/USD	-  15
					,`n EUR/USD	-  11
					,`n GBP/USD	-  5
					,`n NZD/USD	-  18

		}

	If !toggleList
		{
			msgbox, 262144, Currency Pairs, Average Daily Range: `n
   ............................................`n 1.	 EUR/NZD
					,`n 2.	 GBP/JPY 
					,`n 3.	 GBP/AUD 
					,`n 4.	 GBP/CAD
					,`n 5.	 GBP/USD
					,`n 6.	 EUR/JPY
					,`n 7.	 EUR/AUD
					,`n 8.	 GBP/CHF
					,`n 9.	 EUR/CAD
					,`n 10.	 AUD/JPY
					,`n 11.	 EUR/USD
					,`n 12.	 CAD/JPY
					,`n 13.	 USD/CAD
					,`n 14.	 NZD/JPY
					,`n 15.	 AUD/USD
					,`n 16.	 CHF/JPY
					,`n 17.	 USD/CHF
					,`n 18.	 NZD/USD
					,`n 19.	 USD/JPY
					,`n 20.	 CAD/CHF
					,`n 21.	 AUD/CHF
					,`n 22.	 AUD/CAD
					,`n 23.	 EUR/GBP
					,`n 24.	 EUR/CHF
		}	
	return

;Display keyboard hotkeys

	:?*:hh::
	keywait, h
	msgbox,	262144, Assistant, MT4 k-hotkeys	~ all charts	• in subset`n
			. =================================
			.`n cc 		- open calculator
			.`n ee 		- show all currency pairs
			.`n dd 		- subset of currency pairs
			.`n ff 		- fullscreen toggle
			.`n hh 		- show all hot keys
			.`n ll		- view hotkeys in use
			.`n nn		- open new order window
			.`n mm		- view all mouse hotkeys
			.`n ss		- show pairs spread & price
			.`n tt 		- show all template setups
			.`n ww 		- setup windows for scrolling
			.`n`n Tab  e 		~ next timeframe
			.`n Tab  w 		- reset & setup windows
			.`n Tab  q 		~ previous timeframe
			.`n`n Tab  d 		- next timeframe
			.`n Tab  s 		~ input timeframe		
			.`n Tab  a 		- previous timeframe
			.`n`n Tab  z 		- close all sub-windows
			.`n`n Capslock  e 	~ next template 
			.`n Capslock  w	~ input template 
			.`n Capslock  q	~ previous template
			.`n`n Capslock  d 	- next template
			.`n Capslock  s	- next chart window
			.`n Capslock  a	- previous template
			.`n`n Capslock  g 	- next profile
			.`n Capslock  f	- previous profile
			.`n`n Capslock  c 	- clone current window
			.`n Capslock  z	- closes current window
			.`n`n Shift  d 		• next currency pair
			.`n Shift  s 		- open all symbols (24)
			.`n Shift  a 		• previous currency pair
			.`n`n Shift  c 		- next currency pair
			.`n Shift  x 		- input currency pair
			.`n Shift  z 		- previous currency pair
			.`n`n Control		- toggles QuickSilver & Wave .
	return

;Display mouse hotkeys

	:?*:mm::
	keywait, d
	msgbox, 262144, Assistant, MT4 m-hotkeys  	~~>	description`n
			 =================================
			.`n Ctrl  & lbutton  	~~>	open trendline
			.`n Ctrl  & rbutton  	~~>	enter trade
			.`n`n Ctrl & wheelup  	~~>	horizontal line
			.`n Ctrl & wheeldn	~~>	vertical line
			.`n Shift & wheeldn	~~>	trend line
			.`n`n wheelup twice 	~~> 	crosshair
			.`n wheeldn thrice	~~>  	tooltip
			.`n`n xbutton1 short  	~~>	single pip range
			.`n xbutton2 short  	~~>	multi  pip range
			.`n`n xbutton1 long  	~~>	up arrow
			.`n xbutton2 long  	~~>	down arrow
			.`n`n mbutton chart  	~~>	capture image+
			.`n mbutton order  	~~>	calculate spread
			.`n mbutton edges 	~~>	open/toggle MT4 & Journal
			.`n`n ^mbutton edges	~~>	save detailed report from MT4
			.`n +mbutton edges 	~~>	check & compare spread of current pair
			.`n`n F1 , F2 , F3 	~~>	calc projected profit from bars chosen .

	return

;Display currency price and spread

	:?*:ss::
	keywait, s
	Gosub, JournalData
	if CkOpen = closed
	 return
	msgbox,,Spread & Price of Currency Pairs, %JournalData%
	return

;Insert horizontal, vertical, and trend- lines

	^wheelup::
	keywait, ctrl
	IfWinActive ahk_class MetaQuotes::MetaTrader::4.00
	 {
	  MouseGetPos, xl, yl
	  if (xl <= 1862) and (yl > 88) and (yl < 990)
	   send, {alt}{i}{l}{v}{click}
	 }
	return

	^wheeldown::
	keywait, ctrl
	IfWinActive ahk_class MetaQuotes::MetaTrader::4.00
	 {
	  MouseGetPos, xl, yl
	  if (xl <= 1862) and (yl > 88) and (yl < 990)
	   send, {alt}{i}{l}{h}{click}
	 }
	return

	+wheeldown::
	keywait, ctrl
	IfWinActive ahk_class MetaQuotes::MetaTrader::4.00
	 {
	  MouseGetPos, xl, yl
	  if (xl <= 1862) and (yl > 88) and (yl < 990)
	   send, {alt}{i}{l}{t}
	 }
	return

;Call excel macro to enter paste trade

	^rbutton::
	Gosub, Title
	IfWinExist %JournalTitle% ahk_class XLMAIN
	 WinActivate
	send, ^+t
	return

;Scroll using wheelup/down

	!wheeldown::wheeldown
	return
	
	!wheelup::wheelup
	return

;wheeldown/up for tooltip/crosshair or mouseclicks

	wheelup::
	MouseGetPos, xp, yp
	Gosub, crossclick
	return

	wheeldown::
	MouseGetPos, xd, yd
	Gosub, tooltip
	return

;toggle crosshair and alternatively send a mouseclick

	crossclick:
	if wUpScroll > 0
	 {
	  wUpScroll += 1
	  return
	 }
	wUpScroll = 1
	SetTimer, wpCount, 220 ;Wait for more presses within a 220 millisecond window
	return

	wpCount:
	SetTimer, wpCount, off

	if wUpScroll = 1
	 {
	  if (xp > 1050 and xp < 1110) and (yp > 998 and yp < 1050)
	   {
	    Gosub, :?*:ww					;setup charts by scroll
	   }
	  if (xp < 1050 or xp > 1110) and (yp > 998 and yp < 1050)
	   {
	    Gosub, Capslock & a					;previous template
	   }
	  if (yd < 996 or yd > 1048)
	   {
	    click right
	   }
	 }
	if (wUpScroll > 1 and wUpScroll < 4) 
	 {
	  click 2
	 }
	if (wUpScroll > 3) 
	 {
	  ch := !ch
	  if ch
	   send, ^f
	  if !ch
	   click
	 }

	;reset the count to prepare for the next series of presses
	wUpScroll = 0

	return

;ToolTip_DayHour called based on mouse movement

	tooltip:
	if wDownScroll > 0
	 {
	  wDownScroll += 1
	  return
	 }
	wDownScroll = 1
	SetTimer, wdCount, 220 ;Wait for more presses within a 220 millisecond window
	return

	wdCount:
	SetTimer, wdCount, off
	if wDownScroll = 1 ; The key was pressed once.
	 {
	  if (xd > 1050 and xd < 1110) and (yd > 998 and yd < 1050)
	   {
	    Gosub, Capslock & s					;goto next chart window
	   }
	  if (xd < 1050 or xd > 1110) and (yd > 998 and yd < 1050)
	   {
	    Gosub, Capslock & d					;next template
	   }
	  if (yd < 996 or yd > 1048)
	   {
	    click
	   }
	 }
	if (wDownScroll > 1 and wDownScroll < 5) 		;click and hold
	 {
	  click down
	 }
	if (wDownScroll > 4)
	 {
	  if (xd >= 1862) or (yd < 88) or (yd > 990)
	   {
	    Send, {vkADsc120} ;for function keys		;toggles volume on and off
	   }
	  if (xd <= 1860) and (yd > 100) and (yd < 980)
	   {
	    wDS := !wDS
	    if !wDS	;then cancel tooltip
	     tooltip
	    if wDS	;then activate tooltip
	     {
	      While (wDownScroll > 4 and wDownScroll < 8)	;date, time, and pip range balloon
	       {
		MouseGetPos, xpos1, ypos1
		Sleep, 10
		MouseGetPos, xpos2, ypos2

		if (xpos1 < 10) or (xpos2 < 10) or (xpos1 > 1860) or (xpos2 > 1860) or (ypos1 < 100) or (ypos1 > 980) or (ypos2 < 100) or (ypos2 > 980)
		 tooltip
		else if (xpos1 <> xpos2) or (ypos1 <> ypos2)
		 Gosub, ToolTip_DayHour
	       }
	     }
	   }
	 }
	;reset the count to prepare for the next series of presses
	wDownScroll = 0
	return

;Places ToolTip for day and hour depending on timeframe
	ToolTip_DayHour:
	IfWinActive ahk_class MetaQuotes::MetaTrader::4.00	
	 {
	  Gosub, DateTime 

	  ;in order to get southern california time
	  ;1 = Monthly 2 = Weekly 3 = Daily 4 = H4
	  ;5 = H1      6 = M30    7 = M15   8 = M5  9 = M1

	  FormatTime, CheckSat, %DayHour%, WDay
	  if (i = 1 and CheckSat = 7)
	    EnvAdd, DayHour, 1, Days
	  if (i > 3)
	    EnvAdd, DayHour, -10, Hours

	;format for the timeframes
	  if (i < 3)
	    FormatTime, DayHour, %DayHour%, MMM d, yyyy	  

	  if (i = 3)
	    FormatTime, DayHour, %DayHour%, MMM d, ddd

	  if (i = 4 or i = 5)
	    FormatTime, DayHour, %DayHour%, ddd h tt

	  if (i > 5)
	    FormatTime, DayHour, %DayHour%, ddd h:mm tt

	  StringReplace, DayHour, DayHour, PM, pm
	  StringReplace, DayHour, DayHour, AM, am

	  Gosub, PriceData

	 ;Determine Pip Range
	  SetFormat FloatFast, 0
	  UpSwing := (H - O)
	  DnSwing := (O - L)
	  if (UpSwing > DnSwing)
	     {
	      pRange := UpSwing
	      arrow   = ↑
	     }
	  if (DnSwing > UpSwing)
	     {
	      pRange := DnSwing
	      arrow   = ↓
	     }
	  pRange := "Range: " pRange arrow

	  if (StrLen(O) > 0)
	      ToolTip, %DayHour%`n%pRange%, xpos2 + 20, ypos2 + 20
	  else
	      ToolTip
	 }

	IfWinNotActive ahk_class MetaQuotes::MetaTrader::4.00
	   {
	    ToolTip
	   }
	return

;xbutton for journal or volume

	xbutton1::
	MouseGetPos, bx1, by1
	if (bx1 >= 1862) or (by1 < 86 or by1 > 998)
	 Gosub, xb1Up					;raise volume
	else
	 Gosub, RangeArrow1				;insert arrow / acquire ranges
	return

	xbutton2::
	MouseGetPos, bx2, by2
	if (bx2 >= 1862) or (by2 < 86 or by2 > 998)
	 Gosub, xb2Down					;lower volume
	else
	 Gosub, RangeArrow2				;insert arrow / acquire ranges
	return

;raise/lower volume
	
	xb1Up:
	vUp := 1
	While vUp = 1
	 {
	  Send, {vkAFsc130}
	  sleep, 60
	  vUp := GetKeyState("xbutton1","P")
	 }
	return

	xb2Down:
	vDn := 1
	While vDn = 1
	 {
	  Send, {vkAEsc12E}
	  sleep, 60
	  vDn := GetKeyState("xbutton2","P")
	 }
	return

;(1) acquire single or multi candle pip ranges
;(2) insert top/bottom arrows

	RangeArrow1:
	sleep, 400

	xb1 := GetKeyState("xbutton1","P")
	if xb1 = 0
	 Gosub, sRange				;(1) calc range of single session
	else
	 send, {alt}{i}{r}{a}{click}		;(2) insert up arrow
	return

	RangeArrow2:
	sleep, 400

	xb2 := GetKeyState("xbutton2","P")
	if xb2 = 0
	 Gosub, mRange				;(1) calc range of multi session
	else
	 send, {alt}{i}{r}{d}{click}		;(2) insert down arrow
	return

;Calculate for 1 candle pip range from high, low, open, close

	sRange:
	;click					;to get rid of crosshair or tooltip
	tooltip

	Gosub, DateTime
	Gosub, PriceData

	SetFormat FloatFast, 0.1

	OpentoClose 	:= Abs(O - C)
	OpentoLow 	:= Abs(O - L)
	OpentoHigh 	:= Abs(H - O)
	HightoClose 	:= Abs(H - C)
	ClosetoLow  	:= Abs(C - L)
	PipRange 	:= Abs(H - L)

	msgbox,  262144, Pip Ranges, %DateTime%`n-----------------------------`n Open to Close:`t %OpentoClose% `n`n Open to High: `t %OpentoHigh% `n Open to Low: `t %OpentoLow% `n`n Close to High: `t %HightoClose% `n Close to Low: `t %ClosetoLow% `n`n Pip Range: `t %PipRange% `n -----------------------------, 10

	return

;Calculate for 2 candle pip range from high, low, open, close and the length of the trade

	mRange:
	;click			;to get rid of crosshair or tooltip
	tooltip

	KeyWait, xbutton2, U
	  ck := 0		;first click of xbutton2

	Gosub, PriceData

	O1 := O
	H1 := H
	L1 := L
	C1 := C
	
	Gosub, DateTime
	DT1 := DateTime
	DH1 := DayHour
	
	KeyWait, xbutton2, D T5	;timeout of 5 secs = errorlevel 1
	if ErrorLevel
	 {
	  return
	 }
	else
	 {
	  ck := 1		;second click of xbutton2	
	 }

	Gosub, PriceData	

	O2 := O
	H2 := H
	L2 := L
	C2 := C

	SetFormat FloatFast, 0.1

	OO := Abs(O1 - O2)
	OL := Abs(O1 - L2)
	OC := Abs(O1 - C2)
	OH := Abs(O1 - H2)
	HL := Abs(H1 - L2)
	LH := Abs(L1 - H2)
	
	Gosub, DateTime
	DT2 := DateTime
	DH2 := DayHour
	DH3 := DH1

	;subtract start date and end date
	EnvSub, DH3, %DH2%, seconds
	EnvSub, DH2, %DH1%, seconds

	;Get positive difference from dates
	if DH2 > 0
	   DHT := DH2
	if DH3 > 0
	   DHT := DH3

	Gosub, TimeLength

	if OO = 0
	   GoSub, sRange
	if OO <> 0
	   msgbox,  262144, Pip Ranges, %DT1%`n-----------------------------`n Open to Open:`t %OO% `n`n Open to High: `t %OH% `n Open to Close: `t %OC% `n Open to Low: `t %OL% `n`n High to Low: `t %HL% `n Low to High: `t %LH%`n-----------------------------`n%DT2%`n-----------------------------`nLength of trade:`n%TimeDifference%`n-----------------------------`n

	return

;Get Date and Time

	DateTime:
	StatusBarGetText, DateTime, 3, A

	;removes unnecessary characters to conform to required pattern
	StringReplace, DateTime, DateTime, %A_SPACE%,, All
	StringReplace, DateTime, DateTime, .,, All
	StringReplace, DateTime, DateTime, :,, All

	Gosub, TimeFrame				;acquire i variable

	DayHour := DateTime				;used if called by other scripts before format

	if (ck = 1)
	 {
	  transfer := TimeArray[i]
	  EnvAdd, DateTime, %transfer%, seconds		;add period to get accurate end date
	  ck := 0					;reset ck
	 }

	FormatTime, CheckSat, %DateTime%, WDay
	if (i = 1 and CheckSat = 7)			;check if monthly chart has a Sat date & change to Sun
	 EnvAdd, DateTime, 1, Daysf
	if (i > 3)
	 EnvAdd, DateTime, -10, Hours			;adjust to pacific time

	;format times according to timeframe
	if (i < 3)
	  FormatTime, DateTime, %DateTime%, ddd, MMM d, yyyy

	if (i = 3)
	  FormatTime, DateTime, %DateTime%, ddd, MMM d

	if (i = 4 or i = 5)
	  FormatTime, DateTime, %DateTime%, MMM d, ddd h tt

	if (i > 5)
	  FormatTime, DateTime, %DateTime%, MMM d, ddd h:mm tt

	StringReplace, DateTime, DateTime, PM, pm
	StringReplace, DateTime, DateTime, AM, am
	return

;Convert seconds to length of time (i.e. years, months, weeks, days, minutes)

	;#Persistent
	TimeLength:

	PeriodArray 	:= {1: "year", 2: "month", 3:"week", 4: "day", 5: "hour", 6: "min"}
	DivisorArray	:= {1: 31557875.59, 2: 2629822.9658, 3: 604800.0, 4: 86400.0, 5: 3600.0, 6: 60.0}

	;seconds to add to time in order to get accurate time span of trade [instead of calculating from start of session to start of final session it will calc start session to end of final session]
	;			month		 week	      1 day	  4 hrs	       1 hr	  30 m	     15 m      5m	1m
	TimeArray	:= {1: 2629822.9658, 2: 604800.0, 3: 86400.0, 4: 14400.0, 5: 3600.0, 6: 1800.0, 7: 900.0, 8: 300.0, 9: 60}

	DHT := DHT + TimeArray[i]						;add period to get accurate time span of trade
	TimeDifference :=
	TF :=
	i  :=

	While i < 6								;loop through all time frames
	 {
	  i++
	  TF := SubStr(DHT/DivisorArray[i], 1, 1)				;check the left most digit
	  if TF <> 0								;if it is a not 0 we have the first time period
	   {
	    Int := floor(DHT / DivisorArray[i])					;get number left of the decimal
	    if Int > 1 								;if greater than 1 add s to label to make it plural
	      PeriodArray[i] := PeriodArray[i] "s"
	    PeriodLabel%i% := Int " " PeriodArray[i] 				;put number left of the decimal with appropriate label
	    TimeDifference := TimeDifference "  " PeriodLabel%i%		;string all periods and labels together
	    DHT -= (floor(DHT / DivisorArray[i]) * DivisorArray[i]) 		;subtract time found from remaining secs
	    TF := 0								;reset to check for next time period
	   } 
	 }

	;Erase Array
	i :=
	loop, 6
	 {
	  i++
	  PeriodLabel%i% :=
	 }
	return

;Determine time frame of current chart

	TimeFrame:
	WinGetTitle, Txt, A
	TimeFrameArray := {1: "Monthly",2: "Weekly",3: "Daily",4: "H4",5: "H1",6: "M30",7: "M15",8: "M5",9: "M1"}
	TFArray := {1: "Monthly",2: "Weekly",3: "Daily",4: "4 Hour",5: "Hourly",6: "30 min",7: "15 min",8: "5 min",9: "1 min"}
	i := 0

	Loop, 9
	{
	  i++
	  TF := InStr(Txt, TimeFrameArray[i])
	  If TF <> 0
	  break
	}
	
	TimeFrame  := TFArray[i]
	TimeFrameN := i

	return
		
;Get Price Data and return it as number with appropriate decimals

	PriceData:
	StatusBarGetText, OpenPrice,	4, A
	StatusBarGetText, HighPrice,	5, A
	StatusBarGetText, LowPrice,	6, A
	StatusBarGetText, ClosePrice,	7, A

	StringTrimLeft, O, OpenPrice,	3
	StringTrimLeft, H, HighPrice,	3
	StringTrimLeft, L, LowPrice,	3
	StringTrimLeft, C, ClosePrice,	3

	StringTrimLeft, OO, OpenPrice,	3
	StringTrimLeft, HH, HighPrice,	3
	StringTrimLeft, LL, LowPrice,	3
	StringTrimLeft, CC, ClosePrice,	3

	Position := InStr(O, "." )

	if Position = 2
	   {
	    O := O * 10000
	    H := H * 10000
	    L := L * 10000
	    C := C * 10000
	   }

	if Position <> 2
	   {
	    O := O * 100
	    H := H * 100
	    L := L * 100
	    C := C * 100
	   }
	return

;Acquire spread and price from Investment journal & current pair on MT4

	;#Persistent
	JournalData:

	Gosub, Title
	IfWinNotExist, %JournalTitle% ahk_class XLMAIN
	 {
	  msgbox, 262208, Journal Closed, Investment journal is not open.`nJournal data cannot be accessed.
	  CkOpen = closed
    	  return
	 }
	else
	 {
	  CkOpen = open
	 }

	;get current currency pair
	WinGetActiveTitle, cPair
	cPair := SubStr(cPair,29,6)
	if cPair = 
	 cPair := sPair

	IfWinExist, %JournalTitle% ahk_class XLMAIN
	 {
	  Xl := ComObjActive("Excel.Application") ;creates a handle

	  ;Xl.Worksheets("Range").Unprotect	
	  Xl.Worksheets("Range").Range("I23:L47").Copy
	  if errorlevel
	   {
	    msgbox, 262208, Journal Closed, Journal data cannot be accessed.
    	    return
	   }

	  JournalData :=
	  JournalData := clipboard
	  clipboard   :=

	  ;Xl.Worksheets("Range").Protect
	
	  WinActivate, ahk_class MetaQuotes::MetaTrader::4.00

	  ;look for matching pair, then find spread and price
	  Loop, Parse, JournalData, `n
	   {
	    ;msgbox %A_Index% is %A_LoopField%
	    tPair := SubStr(A_LoopField,5,6)
	    if (cPair = tPair)
	     {
	      JournalSpread := SubStr(A_LoopField,1,3)
	      JournalPrice  := SubStr(A_LoopField,16,7)
	      break
	     }
	   }
	 }
	else
	 {
	  msgbox, 262208, Journal Closed, Investment journal is not open.`nThe journal data cannot be retrieved.
	 }

	return

;save detailed report from mt4
;call macro to enter trade in journal

	^mbutton::
	KeyWait, Ctrl
	KeyWait, Shift
	KeyWait, mbutton

	BlockInput, MouseMove

	Gosub, Title				;acquires the title of the excel journal
	Gosub, TimeFrame			;acquires the present time frame
	Gosub, Setup				;acquires the setup of the chart

	clipboard :=				;empty clipboard
	clipboard := "C:\Users\image\Documents\trade\fx\statements\fxcm"
	clipwait  				;wait for the clipboard to contain text

	sleep, 400

	send, ^t				;opens the "Terminal"
	click 220, 1000				;focuses the Terminal on "Account History"
	click right 220, 960			;open menu
	send, {down 6}{enter}			;open "save as" window

/*
	;WinWait, Save As ahk_class #32770,,3
	 if ErrorLevel
	  {
	   BlockInput, MouseMoveOff
	   MsgBox,, Error1, Error1: timed out, 3
	   BlockInput, MouseMoveOff
	   return
	  }
	 else
	  {
	   send, {right}				;goto end of name for this file
	   send, {bs 3}xls				;replace htm extension with xls (excel)
	   click 600, 60				;goto file address
	   send, ^v					;replace the file address
	   sleep, 200
	   send, {enter}				;close window & finalize
	   sleep, 200
	   send, {enter}				;close window & finalize
	   sleep, 200
	   send, {y}					;yes to overwite the existing file
	  }
*/
	sleep, 200
	send, !vt					;close the "Terminal"
	sleep, 200

run C:\Users\image\Documents\trade\fx\statements\fxcm\DetailedStatement.xls

	suspend, on
	WinActivate, %JournalTitle% ahk_class XLMAIN
	suspend, off

	WinWait, Microsoft Excel ahk_class #32770,,3
	 if ErrorLevel
	  {
	   BlockInput, MouseMoveOff
	   MsgBox,, Error2, Error2: timed out, 1
	   return
	  } 
	 else
	  {
	   WinWaitActive, Microsoft Excel ahk_class #32770,,2
	    send, {y}					;yes to opening the detailedstatement.xls

	   send, ^+t					;call Journal_Enter_Trade macro

	   WinWaitActive, New Trade Entry,,2		;yes to entering new trade in trade journal
	    IfWinActive, New Trade Entry
	     send, {y}

	   SetTimer, CloseTradeWin, 400

	   WinWaitActive, Trade Setup?,,4		;enter trade setup number
	    IfWinActive, Trade Setup?
	     send, %Setup% {enter}

	   WinWaitActive, Confirm Location,,2		;yes to confirm selected cell
	    IfWinActive, Confirm Location
	     send, {y}

	   WinWaitActive, Time Frame?,,2		;enter time frame number
	    IfWinActive, Time Frame?
	     send, %TimeFrameN% {enter}

	   SetTimer, CloseTradeWin, Off
	  }
	BlockInput, MouseMoveOff
	return

;responds to MT4 windows about trade already in journal and query regarding next entry

	CloseTradeWin:

	IfWinActive, Trade #
	 send, {enter}

	IfWinActive, Enter Next Trade?
	 send, {y}

	return

;Capture image from trading platform to place in Journal

	mbutton::
	^+m::

	IfWinActive ahk_class MetaQuotes::MetaTrader::4.00
	MouseGetPos, x1, y1

	;if mouse is outside of image area toggle to excel journal
	if (y1 < 98) or (y1 > 634 and y1 < 1050)
	 {
	  keywait, mbutton
	  Gosub, TradeToggle
	  return
	 }

	;if mouse is over taskbar area then do nothing
	if (y1 > 1050)
	 {
	  return
	 }	

	;if mbutton is in down position
	;then the image will be for
	;the Optimal Weekly Trade
	sleep, 300
	SN1 := GetKeyState("mbutton","P")
	SN2 := GetKeyState("m","P")
	If SN1 + SN2 > 0
		SN := "non-trade"
	else
		SN := "trade"

	suspend on
	Gosub, Title
	Gosub, DateTime
	Gosub, TimeFrame
	Gosub, PriceData
	suspend off

	;dimensions of image
	width  = 325
	length = 375

	;edges of currency chart
	TopEdge    := 98
	BottomEdge := 634
	LeftEdge   := 10
	RightEdge  := 1862


	InputBox, pI, Image Position,1 ~~> Center
				  &`n2 ~~> Top
				  &`n3 ~~> Bottom
				  &`n4 ~~> Left
				  &`n5 ~~> Right,,162,190,,,,,Input (1 - 5)
	suspend on
	If pI = Input (1 - 5)
	   {
	    ;Msgbox, 262208, Image Capture Canceled, Canceled.`nImage capture procedure will exit.,2
	    suspend off
	    exit
	    }
	If pI is not digit
	   {
	    Msgbox, 262208, Image Capture Canceled, Input is not a number.`nImage capture procedure will exit.,2
	    suspend off
	    exit
	    }
	If pI not between 1 and 5
	   {
	    Msgbox, 262208, Image Capture Canceled, Input is not between 1 and 5.`nImage capture procedure will exit.,2
	    suspend off
	    exit
	    }

	;x2 and y2 are offsets in order to position trade in image

	;Center the trade
	If pI = 1
	  {
	   x2 := -162.5
	   y2 := -140.625
	   x3 := x1 + x2
	   y3 := y1 + y2
	  }
	;place trade on Top
	If pI = 2
	  {
	   x2 := -162.5
	   y2 := 375
	   x3 := x1 + x2
	   y3 := y1 + y2
	  }
	;place trade on Bottom
	If pI = 3
	  {
	   x2 := -162.5
	   y2 := -375
	   x3 := x1 + x2
	   y3 := y1 + y2	
	  }
	;place trade Left
	If pI = 4
	  {
	   x2 := -120
	   y2 := -187.5
	   x3 := x1 + x2
	   y3 := y1 + y2	
	  }
	;place trade Right
	If pI = 5
	  {
	   x2 := -285
	   y2 := -187.5
	   x3 := x1 + x2
	   y3 := y1 + y2	
	  }

	;make sure edges don't go past
	;boundary of chart
	If (x3 < LeftEdge)
	   x3 := LeftEdge
	If (x3 > RightEdge)
	   x3 := RightEdge - width

	If (y3 < TopEdge)
	   y3 := TopEdge
	If (y3 > BottomEdge)
	   y3 := BottomEdge - length

	If (x3 + width > RightEdge)
	   x3 := RightEdge - width
	If (y3 + length > BottomEdge)
	   y3 := BottomEdge - length

	;actually capture image of trade
	WinActivate, ahk_class MetaQuotes::MetaTrader::4.00
	Gosub, Label
	sleep, 200
	
	BlockInput, MouseMove

	send, !		;toggle between chart on foreground
	sleep, 200	;toggle between chart on foreground
	send, cf	;toggle between chart on foreground
	sleep, 200
	send, {F6}	;start image capture
	sleep, 20
	MouseMove, x3, y3
	click down
	sleep, 20
	MouseMove, width, length, 20, R
	click up	;end image capture
	sleep, 200
	send, !		;toggle between chart on foreground
	sleep, 200	;toggle between chart on foreground
	send, cf	;toggle between chart on foreground

	SetTitleMatchMode, 2

	IfWinExist, %JournalTitle%
	 {
	  WinActivate
	  WinWait, %JournalTitle%,,5
	   if ErrorLevel
	    Gosub, Release
	   else
	    send, ^+i
	  WinWait, Insert Image,,5
	   if ErrorLevel
	    Gosub, Release
	   else
	    send, {y}
	  WinWait, Trade Setup of Image?,,5
	   if ErrorLevel
	    Gosub, Release
	   else
	    send, %SN%{enter}

	  ;if it is a date and an optimal weekly trade image
	  If (StrLen(SN) = 8)
	   {
	    WinWait, Setup of Image?,,5
	     if ErrorLevel
	      Gosub, Release
	     else
	      send, %Setup%{enter}
	    WinWait, Currency Pair of Image?,,5
	     if ErrorLevel
	      Gosub, Release
	     else
	      send, %Symbol%{enter}
	    WinWait, Timeframe of Image?,,5
	     if ErrorLevel
	      Gosub, Release
	     else
	      send, %TimeFrame%{enter}
	    WinWait, Date of Image?,,5
	     if ErrorLevel
	      Gosub, Release
	     else
	      send, %DateTime1%{enter}
	    WinWait, Weekday of Image?,,5
	     if ErrorLevel
	      Gosub, Release
	     else
	      send, %DateTime2%{enter}
	    WinWait, Time of Image?,,5
	     if ErrorLevel
	      Gosub, Release
	     else
	      send, %DateTime3%{enter}
	    WinWait, Piprange of Image?,,5
	     if ErrorLevel
	      Gosub, Release
	     else
	      send, %PipRange%{enter}
	    WinWait, Direction of Image?,,5
	     if ErrorLevel
	      Gosub, Release
	     else
	      send, %Direction%{enter}
	   }
	 }
	else
	 {
	  BlockInput, MouseMoveOff
	  msgbox, 262208, Journal Closed, Investment Journal is not open.`nOnce open enter image manually.
	 }

	BlockInput, MouseMoveOff
	suspend off
	return

;Add Labels to ScreenCapture Image on MT4

	;#Persistent
	Label:
	suspend, on

	;Symbol
	WinGetTitle, Symbol
	chr := "["
	position := InStr(Symbol, chr) + 1
	Symbol := SubStr(Symbol,position,6)

	;Direction of trade
	MsgBox, 35, Direction of Trade, Long / buy   =   yes `nShort / sell   =    no  (default), 6

	IfMsgBox, Yes
	  Direction = long
	IfMsgBox, No
	  Direction = short
	IfMsgBox, Timeout	;treat the same as No
	  Direction = short
	IfMsgBox, Cancel
	 {
	  Msgbox, 262208, Image Capture Canceled, Image capture procedure will exit., 2
	  suspend off
	  exit
	 }

	;Current Date/Time
	MsgBox, 35,Date & Time, Current local time  = yes `nMT4 time (default) = no, 2

	IfMsgBox, Yes
	 {
	  FormatTime, DateTime1,, MMMM d, yyyy
	  FormatTime, DateTime2,, ddd h:mm tt
	  FormatTime, DateTime3,, MMM d
	 }
	IfMsgBox, No
	 {
	  FormatTime, DateTime1, %DayHour%, MMMM d, yyyy
	  FormatTime, DateTime2, %DayHour%, ddd h:mm tt
	  FormatTime, DateTime3, %DayHour%, MMM d
	 }
	IfMsgBox, Timeout	;treat the same as No
	 {
	  FormatTime, DateTime1, %DayHour%, MMMM d, yyyy
	  FormatTime, DateTime2, %DayHour%, ddd h:mm tt
	  FormatTime, DateTime3, %DayHour%, MMM d
	 }
	IfMsgBox, Cancel
	 {
	  Msgbox, 262208, Image Capture Canceled, Image capture procedure will exit., 2
	  suspend off
	  exit
	 }
	StringReplace, DateTime2, DateTime2, PM, pm
	StringReplace, DateTime2, DateTime2, AM, am

	;Current Chart TimeFrame
	TimeFrame = %TimeFrame%, %Datetime3%

	;Determine Pip Range
	SetFormat FloatFast, 0
	if (Direction = "Long")
	    PipRange := (H - O)
	if (Direction = "Short")
	    PipRange := (O - L)
	PipRange = %PipRange% pips

	BlockInput, MouseMove

	;Setup
	send, ^b			;opens objects window in MT4 and click on top of list
	WinWaitActive, Objects, , 2
	 if ErrorLevel
	  {
	   suspend, off
	   BlockInput, MouseMoveOff
    	   MsgBox, 262160, Image Capture Canceled, Objects List window is not responding, 4
	   exit
	  }
	 else
	  {
	   mousemove, 170, 70
	   click 2
	   sleep 200
	  }
	WinWaitActive, Label, , 2
	 if ErrorLevel
	  {
	   suspend, off
	   BlockInput, MouseMoveOff
    	   MsgBox, 262160, Image Capture Canceled, Label window is not responding, 4
	   exit
	  }
	 else
	  {
	   ;make sure correct tab (Common) is open in Label window
	   WinGetText, Ldata, Label
	   Loop, Parse, Ldata, `n, `r
	    {
	     temp := A_LoopField
	     break
	    }
	   if (temp = "Visualization")
	    {
	     send, ^{Tab}
	    }
	   else if (temp = "Parameters")
	    {
	     send, +^{Tab}
	    }
	   else if (temp != "Common")
	    {
	     suspend, off
	     BlockInput, MouseMoveOff
    	     MsgBox, 262160, Image Capture Canceled, Common tab in Label window not found., 4
	     exit
	    }
	  }

	;copy Setup
	clipboard =	;empty clipboard
	send, ^c
	sleep 20
	Setup = %clipboard%
	clipboard =	;empty clipboard

	If SN = trade
	 {
	  SN := Setup
	  If SN = QuickSilver
		  SN := 1
	  If SN = Propulsion
		  SN := 2
	  If SN = 5 Bullets
		  SN := 3
	  If SN = Triple Threat
		  SN := 4
	  If SN = Spring MA
		  SN := 5
	  If SN = NTP Scale
		  SN = 6
	  If SN = Candle Pattern
		  SN := 7
	  If SN = PTP Setup
		  SN := 8
	  If SN = Pivot (d)
		  SN := 9
	  If SN = Pivot (w)
		  SN := 9
	  If SN = Pivot (m)
		  SN := 9
	  If SN = Convergence (a)
		  SN := 10
	  If SN = Convergence (b)
		  SN := 10
	  If SN = Alligator Volume
		  SN := 11
	  If SN = Wave
		  SN := 12
	 }
	If SN = non-trade
	 {
	  FormatTime, SN, %DayHour%, MM/dd/yy
	 }
	
	sleep 400 ;allow program time to avoid glitches	
	;goto 'Parameters' tab to input new coordinates
	send, ^{Tab}

	;make sure correct tab (Parameters) is open in Label window
	WinGetText, Ldata, Label
	Loop, Parse, Ldata, `n, `r
	 {
	  temp := A_LoopField
	  break
	 }
	if (temp = "Common")
	 {
	  send, ^{Tab}
	 }
	else if (temp = "Visualization")
	 {
	  send, +^{Tab}
	 }
	else if (temp != "Parameters")
	 {
	  suspend, off
	  BlockInput, MouseMoveOff
    	  MsgBox, 262160, Image Capture Canceled, Parameters tab in Label window not found., 4
	  exit
	 }

	;acquire new coordinates... 10 (left edge) and 90 (top edge) are offsets for the x/y options of Setup Label
	 x4 := x3 + 8 - 10		;acquire new x coordinate
	 y4 := y3 + 8 - 90		;acquire new y coordinate

	;insert new coordinates
	send, {tab}			;go to x coordinate
	send, %x4%{tab}			;new x coordinates to move setup label
	send, %y4%{tab}			;new y coordinates to move setup label
	send, +^{Tab}			;go back to 'Common' tab
	send, {enter}{esc}		;exit label window
	sleep 200

	;set start position for rest of labels under template label
	mousemove x3 + 12, y3 + 19

	i := 0
	Loop 4
	{
	i += 1
	send, {enter}
	sleep, 20
	send, {esc}
	sleep, 20

	send, {alt}{i}{b}
	sleep, 20
	mousemove 0, 14, 2, R
	sleep, 20
	click
	sleep, 20
	send, %i%
	send, {tab}
	sleep, 20

	if i = 1 
	   send, %Symbol%
	if i = 2 
	   send, %TimeFrame%
	if i = 3 
	   send, %DateTime2%
	if i = 4
	   send, %PipRange%
	
	sleep, 100
	if i = 1 		;sets up size and color defaults
	 {
	  send, {tab 2}
	  send, 8		;font size
	  sleep, 20
	  send, {tab}
	  send, {del 10}
	  send, silver		;font color
	  sleep, 20
	 }
	send, {enter}

	if i = 4		;moves cursor away from label to avoid info balloon in image capture
	 {
	  sleep, 20
	  mousemove 0, 14, 2, R
	  sleep, 20
	 }
	}

	;prepare TimeFrame for excel by removing month and day
	Position := InStr(TimeFrame, ",") - 1
	TimeFrame := SubStr(TimeFrame, 1, Position)

	;prepare DateTime for excel by splitting day and time
	FormatTime, DateTime2,%DayHour%, dddd
	FormatTime, DateTime3,%DayHour%, h:mm tt
	StringLower, DateTime3, DateTime3

	;prepare PipRange for excel by removing "pips" from it
	PipRange := SubStr(PipRange, 1, -5)

	BlockInput, MouseMoveOff
	suspend, off
	return

;acquire Setup

	Setup:

	send, ^b			;opens objects window in MT4
	sleep, 200
	mousemove 210, 80
	click 2				;opens the labels window
	sleep, 200

	clipboard :=			;empty clipboard
	send ^c				;copy the already selected label
	sleep, 200
	setup = %clipboard%		;transfer clipboard
	clipboard :=			;empty clipboard

	send, {esc 2}			;close objects window & label window

	if setup = QuickSilver
		setup := 1
	if setup = Propulsion
		setup := 2
	if setup = 5 Bullets
		setup := 3
	if setup = Triple Threat
		setup := 4
	if setup = Spring MA
		setup := 5
	if setup = NTP Scale
		setup = 6
	if setup = Candle Pattern
		setup := 7
	if setup = PTP Setup
		setup := 8
	if setup = Pivot (d)
		setup := 9
	if setup = Pivot (w)
		setup := 9
	if setup = Pivot (m)
		setup := 9
	if setup = Convergence (a)
		setup := 10
	if setup = Convergence (b)
		setup := 10
	if setup = Alligator Volume
		setup := 11
	if setup = Wave
		setup := 12
	return

;calc projected profit % of pips captured from bars chosen
;F1 starts script, F2 chooses bar, F3 ends script - Project.mq4
;F1 starts script, F2 chooses bar, F3 ends script - Swing.mq4
#IfWinActive ahk_class MetaQuotes::MetaTrader::4.00
#Persistent

	SC03B::							;F1 button

	  a := "Attach the Project or Swing projection?"
	  b := "Project = yes    ||    Swing = no "
	  c := StrLen(a) + 12

	  a := PadStr(a, c, "R")
	  b := PadStr(b, c, "R")

	msgbox, 3, Projection, %a% `n %b%, 10
	IfMsgBox, Yes
	 Gosub, Project
	IfMsgBox, No
	 Gosub, Swing
	return

	Project:
	;remove any previous instance of Project indicator in MT4
	send, +Q

	;acquire currency pair
	WinGetTitle, pair
	position := InStr(pair, "[") + 1
	pair     := SubStr(pair, position, 6)

	;acquire spread
	FilePath := "C:\Users\image\AppData\Roaming\MetaQuotes\Terminal\BB190E062770E27C3E79391AB0D1A117\MQL4\Files\Spread.txt"
	FileRead, TxtFile, %FilePath%
	spread := SubStr(TxtFile, 1, 3)
	symbol := SubStr(TxtFile,InStr(TxtFile,"(") + 1, 6)

	if (symbol != pair)
	 Loop, 10
	 {
	  FileRead, TxtFile, %FilePath%
	  spread := SubStr(TxtFile, 1, 3)
	  symbol := SubStr(TxtFile,InStr(TxtFile,"(") + 1, 6)

	  iteration := A_Index + 90
	  Progress, %iteration%, , Calculating..., %pair% Spread
	  
	  if (GetKeyState("Esc", "P") = 1)			;Esc button pressed down
	   {
	    Keywait, Esc					;wait for Esc to be released
	    break
	   }

	  if (symbol = pair)
	   break
	  else
	   sleep 400
	 }

	Progress, Off
	TxtFile = 						;free memory

	if (symbol != pair)
	 {
	  a := "          "
	  b := "                              "
	  msgbox, 262148, Project End, %a% The currency pairs do not match.`n     (active window symbol & spread symbol)`n`n %b% Continue?
	  IfMsgBox, No
	   exit
	  IfMsgBox, Yes
	   spread := 2
	 }

	WinActivate, ahk_class MetaQuotes::MetaTrader::4.00
	WinWaitActive, ahk_class MetaQuotes::MetaTrader::4.00,, 4
	IfWinActive, ahk_class MetaQuotes::MetaTrader::4.00

	IfWinActive ahk_class MetaQuotes::MetaTrader::4.00
	 {
	  Gui, New
	  Gui, Color, White
	  Gui, +AlwaysOnTop
	  Gui, +ToolWindow +Caption +Border
	  Gui, Add, Text,, Starting Equity:
	  Gui, Add, Text,, StopLoss:
	  Gui, Add, Text,, Spread:
	  Gui, Add, Text,, Captured Pips (`%):
	  Gui, Add, Text,, Risk (`% of equity):
	  Gui, Add, Text,, Use complete labels:
	  Gui, Add, Edit, vequity number w40 center ym, 2000  	;the ym option starts a new column of controls
	  Gui, Add, Edit, vstoploss number w40 limit3 center, 60
	  Gui, Add, Edit, vspread w40 limit3 center, %spread%
	  Gui, Add, Edit, vpercent w40 number limit2 center, 80
	  Gui, Add, Edit, vrisk number w40 limit2 center, 3
	  Gui, Add, Checkbox, vlabel  w40 -wrap checked
	  Gui, Add, Button, x58 y166 center w60 +default, Start
	  Gui, Show, center, Project Inputs
	  return
	 }

	ButtonStart:
	Gui, Submit						;save the input from the user to each control's associated variable.
	IfWinNotActive, ahk_class MetaQuotes::MetaTrader::4.00
	 exit
	risk 	:= risk / 100
	percent := percent /100
	suspend, on

	send, ^P						;call Project indicator in MT4 - inserts accurate arrows
	WinWaitActive, Custom Indicator - Project,, 2
	if ErrorLevel
	 {
	  a := "Project indicator did not load."
	  b := "Projection Series terminated."
	  c := StrLen(a) + 6

	  a := PadStr(a, c, "R")
	  b := PadStr(b, c, "R")

	  Msgbox, 0, Project, %b% `n %a%, 3
	  exit
	 }

	IfWinActive, Custom Indicator - Project
	 send, {enter}	

	sleep, 200

	FilePath := "C:\Users\image\AppData\Roaming\MetaQuotes\Terminal\BB190E062770E27C3E79391AB0D1A117\MQL4\Files\Project.txt"
	FileDelete, %FilePath%
	FileAppend,
	 (
	  risk     = %risk%
	  label    = %label%
	  equity   = %equity%
	  percent  = %percent%
	  stoploss = %stoploss%`n
	 ), %FilePath%

	suspend, off
	Gosub, sProject
	return

	Swing:
	;remove any previous instance of Swing indicator in MT4
	send, +Q

	;acquire spread
	FilePath := "C:\Users\image\AppData\Roaming\MetaQuotes\Terminal\BB190E062770E27C3E79391AB0D1A117\MQL4\Files\Spread.txt"
	FileRead, TxtFile, %FilePath%
	spread := SubStr(TxtFile, 1, 3)

	WinActivate, ahk_class MetaQuotes::MetaTrader::4.00
	WinWaitActive, ahk_class MetaQuotes::MetaTrader::4.00,, 4
	IfWinActive, ahk_class MetaQuotes::MetaTrader::4.00

	IfWinActive ahk_class MetaQuotes::MetaTrader::4.00
	 {
	  Gui, New
	  Gui, Color, White
	  Gui, +AlwaysOnTop
	  Gui, +ToolWindow +Caption +Border
	  Gui, Add, Text,, Starting Equity:
	  Gui, Add, Text,, StopLoss:
	  Gui, Add, Text,, Spread:
	  Gui, Add, Text,, Captured Pips (`%):
	  Gui, Add, Text,, Risk (`% of equity):
	  Gui, Add, Text,, Pullback (pips):
	  Gui, Add, Edit, vequity number w40 center ym, 2000  		;the ym option starts a new column of controls
	  Gui, Add, Edit, vstoploss number w40 limit3 center, 60
	  Gui, Add, Edit, vspread w40 limit3 center, %spread%
	  Gui, Add, Edit, vpercent w40 number limit2 center, 80
	  Gui, Add, Edit, vrisk number w40 limit2 center, 3
	  Gui, Add, Edit, vpullback number w40 limit2 center, 30
	  Gui, Add, Button, x50 y172 center w60 +default, Begin
	  Gui, Show, center, Swing Inputs
	  return
	 }

	ButtonBegin:
	Gui, Submit							;save user input
	IfWinNotActive, ahk_class MetaQuotes::MetaTrader::4.00
	 exit
	risk 	:= risk / 100
	percent := percent /100

	send, ^1							;call Swing indicator in MT4

	WinWaitActive, Custom Indicator - Swing,, 2
	If ErrorLevel
	 {
	  a := "Swing indicator did not load."
	  b := "Swing Series terminated."
	  c := StrLen(a) + 6

	  a := PadStr(a, c, "R")
	  b := PadStr(b, c, "R")

	  Msgbox, 0, Swing, %b% `n %a%, 3
	  exit
	 }

	IfWinActive, Custom Indicator - Swing
	 send, {enter}	

	sleep, 200

	FilePath := "C:\Users\image\AppData\Roaming\MetaQuotes\Terminal\BB190E062770E27C3E79391AB0D1A117\MQL4\Files\Swing.txt"
	FileDelete, %FilePath%
	FileAppend,
	 (
	  risk     = %risk%
	  equity   = %equity%
	  percent  = %percent%
	  stoploss = %stoploss%
	  pullback = %pullback%`n
	 ), %FilePath%

	Gosub, sProject
	return

	GuiEscape:
	GuiClose:
	 WinActivate, ahk_class MetaQuotes::MetaTrader::4.00
	 send, +Q						;remove Project indicator from MT4
	 Gui, Destroy
	 exit
	return

	sProject:
	IfWinActive ahk_class MetaQuotes::MetaTrader::4.00
	 SetTimer, RecPos, 40
	return

	SC03C::esc						;avoid MT4 shortcut
	SC03D::esc						;avoid MT4 shortcut

	RecPos:
	IfWinActive ahk_class MetaQuotes::MetaTrader::4.00
	 {			
	  if (GetKeyState("SC03C", "P") = 1)			;F2 button pressed down
	   {
	    KeyWait, SC03C, T1					;wait for F2 to be released
	    if ErrorLevel
	     return
	    else
	      click						;select bar for trade - MT4 Project.mq4
	   }
	  if (GetKeyState("SC03D", "P") = 1)			;F3 button pressed down
	   {
	    KeyWait, SC03D, T1					;wait for F3 to be released
	    if ErrorLevel
	     return
	    else
	     {
	      send, +Q						;tally trade result & exit - MT4 Project.mq4
	      SetTimer, RecPos, off
	      exit
	     }
	   }
	 }
	return

#IfWinActive

	Common:
	;make sure correct tab (Common) is open in Text window
	WinWaitActive, Text,, 4
	WinGetText, Ldata, Text
	Loop, Parse, Ldata, `n, `r
	 {
	  temp := A_LoopField
	  break
	 }
	if (temp = "Visualization")
	 {
	  send, ^{Tab}
	 }
	else if (temp = "Parameters")
	 {
	  send, +^{Tab}
	 }
	else if (temp != "Common")
	 {
	  suspend, off
	  BlockInput, MouseMoveOff
    	  MsgBox, 262160, Project Canceled, temp = %temp% Common tab in Text window not found., 4
	  exit
	 }
	return

	CalTime:
	;in order to get southern california time
	;1 = Monthly 2 = Weekly 3 = Daily 4 = H4
	;5 = H1      6 = M30    7 = M15   8 = M5  9 = M1

	FormatTime, CheckSat, %DayHour%, WDay
	if (i = 1 and CheckSat = 7)
	    EnvAdd, DayHour, 1, Days
	if (i > 3)
	    EnvAdd, DayHour, -10, Hours

	;format for the timeframes
	if (i < 3)
	 {
	  FormatTime, Time1, %DayHour%, MMM d
	  FormatTime, Time2, %DayHour%, yyyy
	 }
	if (i = 3)
	 {
	  FormatTime, Time1, %DayHour%, ddd
	  FormatTime, Time2, %DayHour%, MMM d
	 }
	if (i = 4 or i = 5)
	 {
	  FormatTime, Time1, %DayHour%, ddd
	  FormatTime, Time2, %DayHour%, h tt
	 }
	if (i > 5)
	 {
	  FormatTime, Time1, %DayHour%, ddd
	  FormatTime, Time2, %DayHour%, h:mm tt
	 }

	StringReplace, Time2, Time2, PM, pm
	StringReplace, Time2, Time2, AM, am
	return

#IfWinActive


;function to pad text with space or chosen symbols

	PadStr(str, size, direction="L", char=" ")
	 {
	  strOrig := str
	  if ( direction="L")
	   {
	    loop % size-StrLen(str)
	    str .= char
	   }
	  else If (direction="R")
	   {
	    loop % size-StrLen(str)
	    str := char . str
	   }
	  else If (direction="C")
	   {
	    loop % round(size/2)-(round(StrLen(str)/2))
	    str := char . str
	    loop % round(size/2)-(round(StrLen(strOrig)/2))
	    str .= char
	   }
	  return str
	 }

;function to format currency

	Currency(A)
	{
	 B := floor(A)               	;B is integer
	 stringsplit,C,A,`.        	;C2 is after decimal point
	 Loop Parse,B
   	  {
   	   stringlen,L,B
   	   If (Mod(L-A_Index,3) = 0 and A_Index != L)
       	    x = %x%%A_LoopField%,
	   Else
       	    x = %x%%A_LoopField%
   	  }
	 A := x
	 return A
	 ;msgbox,$%x%;.%C2%          	;result
	}


;Use WheelDown to send numbers 1 - 5
;#IfWinActive, Image Position ;ahk_class #32770

#IfWinActive, ahk_group InputNumber

	WheelDown::

	if keypresses > 0
	 {
	  keypresses += 1
	  return
	 }
	keypresses = 1
	SetTimer, WheelDCount, 1000 ;Wait for more presses within a 1000 millisecond window
	return

	WheelDCount:
	SetTimer, WheelDCount, off
	if keypresses = 1 ; The key was pressed once.
	 {
	  sendinput, 1
	  sleep, 1000
	  send, {enter}
	 }
	else if keypresses = 2
	 {
	  sendinput, 2
	  sleep, 1000
	  send, {enter}
	 }	
	else if keypresses = 3
	 {
	  sendinput, 3
	  sleep, 1000
	  send, {enter}
	 }
	else if keypresses = 4
	 {
	  sendinput, 4
	  sleep, 1000
	  send, {enter}
	 }
	else if keypresses > 4
	 {
	  sendinput, 5
	  sleep, 1000
	  send, {enter}
	 }  
	;reset the count to prepare for the next series of presses
	keypresses = 0
	return

#IfWinActive


;#IfWinActive, Order ahk_class #32770
;Acquire spread of current pair
#Persistent

;	+mbutton::

      Spread:
	sendinput, {F9}
	sleep, 100

      IfWinActive Order ahk_class #32770
       {
	BlockInput, MouseMove

	i := 91
	temp := 0
	SumAverage := 0

	Loop, 10	;acquire spread 10x to average
	 {
	  WinGetText, linefeed
	  Loop, Parse, linefeed, `r
	  {
	   line = %A_Index%

	   if line = 2
	    {
	     temp = %A_LoopField%
	     sPair := SubStr(temp,2,6)
	    }

	   if line = 28
	    {
	     temp = %A_LoopField%

	     n := InStr(temp,"/")
	     n := n + 2

	     temp1 := SubStr(temp,2,7)
	     temp2 := SubStr(temp,n,7)
	    }
  	  }

	  dPos := InStr(temp1, "." )

	  if dPos = 2
	   {
	    temp1 := temp1 * 10000
	    temp2 := temp2 * 10000
	   }

	  if dPos <> 2
	   {
	    temp1 := temp1 * 100
	    temp2 := temp2 * 100
	   }

	  temp := abs(temp1 - temp2)
	  SumAverage := SumAverage + temp

	  suspend, on
	  sleep, 1000
	  suspend, off
	  
	  i += 1

    	  Progress, %i%, , Calculating..., %sPair% Spread
	 }
       }

	Progress, Off

	SetFormat FloatFast, 0.1
	Spread := SumAverage / 10

	;close New Order window
	send, {esc}

	BlockInput, MouseMoveOff

	;MsgBox,, Spread, %sPair% = %Spread%

	return

#IfWinActive
#IfWinActive

;acquire the last 3 days of month excluding Sat
;setup msgbox if needed

	L3DM:

	Loop, 3
	 {
	  i = %A_Index%

	  if wkday = 7
	   {
            rDate%i% := LDM - i
	   }
	  else if wkday > 2
	   {
            t := i - 1
            rDate%i% := LDM - t
	   }
	  else if wkday = 2
	   {
            if i < 3
	     {
	      t := i - 1
	     }
	    else
	     {
	      t := i
	     }
            rDate%i% := LDM - t
	   }
	  else if wkday = 1
	   {
            if i = 1
	     {
	      t := i - 1
	     }
	    else
	     {
	      t := i
	     }
            rDate%i% := LDM - t
	   }
	 }

	if rDays = 1
	 n = day
	if rDays > 1
	 n = days

	Reminder =
	 (
	  Remember to 'save as' a detailed statement from MT4`n  before the end of the month to avoid losing data.`n
	           ~ %rDays% %n% left in %cMnth% ~
	 )

	return

;function acquires the last day and date of the month

	LDOM(TimeStr="")
	 {
  	  If TimeStr =
	     TimeStr = %A_Now%
	  StringLeft Date, TimeStr, 6	; extract YearMonth from given or current date
	  Date1 = %Date%
	  Date1+= 31, D			; change date to one of first few days in next month
	  StringLeft  Date1, Date1, 6	; extract YearMonth of the following month
	  Date1-= %Date%,  D		; difference in days between these months ; last day of month in d format
	  Date2 = %Date%%Date1%		; last date of month in yyyymmdd format
	  Return Date2
	 }