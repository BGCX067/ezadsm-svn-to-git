; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         A.N.Other <myemail@nowhere.com>
;
; Script Function:
;	Template script (you can customize this template by editing "ShellNew\Template.ahk" in your Windows folder)
;
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#NoTrayIcon
SetWorkingDir, C:\Program Files\IBM\ADSM\baclient
SysGet, VirtualScreenWidth, 78
SysGet, VirtualScreenHeight, 79

Gui +LastFound -AlwaysOnTop -Caption +ToolWindow
Gui, Color, FFFFFF
Gui, Font, CBlack
Gui, Font,, Tahoma
Gui, Font, s72
Gui, Add, Text, x-0 w1100 vStatus, Connecting to server...
Gui, Font, s32
Gui, Add, Text, x-0 y+20 w800 hp+50 Wrap vInfo,
Gui, Show, x0 y0 h%VirtualScreenHeight% w%VirtualScreenWidth%
GuiControl,, Info,

DIR := "C:\Program Files\IBM\ADSM\baclient"
CMD := DIR . "\DSMADMC.EXE -console -tcps=bsadsp03 -id=console -passw=console"
IfNotExist, %DIR%
{
  Msgbox, 16, EZADSM, %DIR% not found.
  ExitApp
}
CMDret_Stream(CMD,"",DIR)

CMDret_Output(StrOut, CMDname="")
{ ;DO NOT REMOVE!
  
;Connected to server
IfInString, StrOut, Session established with server BSADSP03
{
  GuiControl,, Status, Waiting for status...
  QueryReq()
}

;Mount request
IfInString, StrOut, Mount 8MM Volume 
  {
    ;6 = volume #
    ;16 = minutes remaining
    StringSplit, MR, StrOut, %A_Space%
    
    ;Get the remaining time first - its easier to set the infoline this way
    MR16hours := MR16 / 60
    StringSplit, MR16hours, MR16hours, .
    MR16mins := Mod(MR16, 60)
    TimeLeft = Time remaining: %MR16hours1% hours %MR16mins% minutes
    InfoText = %TimeLeft%
    
    ;Now determine if its asking for SCRTCH or a specific tape and display
    StatusText = Tape request: %MR6%
    If MR6 is Number
    {
      If StrLen(MR6) > 4
      {
        StringTrimLeft, MR6, MR6, 3
      }
      If SubStr(MR6, 1, 1) = 0
        StringTrimLeft, MR6, MR6, 1
      If SubStr(MR6, 1, 1) = 0
        StringTrimLeft, MR6, MR6, 1
      StatusText = Tape request: %MR6%
      If MR6 = 36
         StatusText = Tape request: %MR6% (Mon)
      If MR6 = 32
        StatusText = Tape request: %MR6% (Tue)
      If MR6 = 30
        StatusText = Tape request: %MR6% (Wed)
      If MR6 = 23
        StatusText = Tape request: %MR6% (Thu)
      If MR6 = 21
        StatusText = Tape request: %MR6% (Fri)
    } 
    If MR6 in SCRTCH,DBBK.1
    {
      StatusText = Tape request: SCRATCH
      ScratchTapes := GetScratchTapes()
      InfoText = %TimeLeft%`nScratch tapes: %ScratchTapes%
    }
    Gui, Color, 000000
    Gui, Font, s72 cRed
    GuiControl,, Status, %StatusText%
    GuiControl, Font, Status
    Gui, Font, s24
    GuiControl,, Info, %InfoText%
    GuiControl, Font, Info
  }

;No outstanding requests
IfInString, StrOut, QUERY REQUEST: No requests are outstanding
{
  Gui, Color, FFFFFF
  Gui, Font, s64 cLime
  GuiControl, Font, Status
  GuiControl,, Status, No outstanding requests
  GuiControl,, Info,
}

;Tape mounted
IfInString, StrOut, ANR8328I
{
  Gui, Color, FFFFFF
  Gui, Font, s64 cBlack
  GuiControl, Font, Status
  GuiControl, Font, Info
  StatusText = Tape %MR6% inserted
  InfoText =
  GuiControl,, Status, %StatusText%
  GuiControl,, Info, 
  ScratchTapes =
  QueryReq()
}

;Tape ejected
IfInString, StrOut, ANR8468I
{
  Gui, Font, s24 cLime
  GuiControl, Font, Info
  StringSplit, EJ, StrOut, %A_Space%
  ;4 = volume
  If StrLen(EJ4) > 4
  {
    StringTrimLeft, EJ4, EJ4, 3
    If SubStr(EJ4, 1, 1) = 0
      StringTrimLeft, EJ4, EJ4, 1
    If SubStr(EJ4, 1, 1) = 0
      StringTrimLeft, EJ4, EJ4, 1
  }
  InfoText = %EJ4% ejected
  GuiControl,, Info, %InfoText%
}

;Mount request cancelled
IfInString, StrOut, ANR1400W
{
  Gui, Color, 000000
  Gui, Font, s64 cWhite
  GuiControl, Font, Status
  GuiControl,, Status, Mount request cancelled
  Gui, Font, s24 cWhite
  GuiControl, Font, Info
  GuiControl,, Info, Waiting for next request...
}

;Server shutdown
IfInString, StrOut, ANR0991I
{
  Gui, Color, FFFFFF
  Gui, Font, s72 cRed
  GuiControl, Font, Status
  GuiControl,, Status, Server shutdown
  SetTimer, RestartCountdown, 1000
}

} ;DO NOT REMOVE!

;SetTimer, LANDeskCheck, 3000
OnExit, CleanUp
Escape::Goto CleanUp
!F10::ListVars
;~ !F11::
;~ GuiControl,, Info, Searching for scratch tapes...
;~ testjz := GetScratchTapes()
;~ GuiControl,, Info, Scratch tapes: %testjz%


Return
;~ #######################
;~ #######################
;~ #######################
;~ ### END OF AUTOEXEC ###
;~ #######################
;~ #######################
;~ #######################

CleanUp:
Process, Close, %cmdretPID%,
ExitApp

;LANDesk auto-accept
;~ LANDeskCheck:
  ;~ IfWinExist, ahk_class #32770
  ;~ {
    ;~ WinActivate
    ;~ Sleep, 1500
    ;~ SendInput, y
  ;~ }
;~ Return

RestartBatch:
  FileAppend,
(
taskkill /F /IM "EZ ADSM.EXE"
"%A_ScriptDir%\%A_ScriptName%"
del C:\ezadsmrestart.bat
), C:\ezadsmrestart.bat
  Run, C:\ezadsmrestart.bat,,Hide
Return

RestartCountdown:
  If ServerShutdown = 1
  {
    If SSSec = 00
    {
      If SSMin = 0
      {
        Goto, RestartBatch
      }
      SSSec := 59
      SSMin--
    }
    Else
    {
;~       SSSec--
      SSSec -= 1
      If Strlen(SSSec) = 1
        SSSec := "0" . SSSec
    }
  }
  Else
  {
    SSMin = 10
    SSSec = 00
    Process, Close, %cmdretPID%,
    ServerShutdown = 1
  }
  Gui, Font, s24
  GuiControl,, Info, Reconnecting in %SSMin%:%SSSec%
Return


QueryReq()
{
  Run, DSMADMC.EXE -tcps=bsadsp03 -id=G018604S -passw=adsm query req,, Hide
}


GetScratchTapes()
{
  RunWait, dsmadmc -tcps=bsadsp03 -id=G018604S -passw=adsm select volume_name from volumes >> vol.txt,,Hide
  Sleep, 1000
  FileRead, vol, vol.txt

  Loop, Parse, vol, `n
  {
      VolRetrieved = 0
      Loop, Parse, A_LoopField, %A_Space%
      {
          If A_LoopField is number
          {
  ;~ 			MsgBox %A_LoopField%
              tape = %A_LoopField%
              If StrLen(tape) > 4
              {
                  StringTrimLeft, tape, tape, 3
              }
              If SubStr(tape, 1, 1) = 0
                  StringTrimLeft, tape, tape, 1
              If SubStr(tape, 1, 1) = 0
                  StringTrimLeft, tape, tape, 1
              If VolumeList
                  VolumeList = %VolumeList%`,%tape%
              Else
                  VolumeList = %tape%
              VolRetrieved = 1
              If A_LoopField = 181
                  VolRetrieved = 2
              Break
          }
          If VolRetrieved = 1
              Continue
          If VolRetrieved = 2
              Break
  }
  If VolRetrieved = 2
              Break
  }

  Sort Volumelist, N D,
  ;~ msgbox %volumelist%

  i = 1
  Loop, Parse, Volumelist, `,
  {
;~       FileAppend, i=%i%%a_tab%%A_Loopfield%`n, C:\test.txt
      If (A_LoopField - i > 1) ;If the sequence jumps by more than 1, add the missing tapes to catch up.
      {
        CatchUp := A_LoopField - i - 1
        Loop %CatchUp%
        {
          MissingTapes := MissingTapes A_Space i + 1
          i++
        }
      }
      If Not i = A_Loopfield
      {
          If i not in 12,13,21,23,30,32,36,199 ;Eliminate daily tapes (and known bad tapes) from the Scratch tape list
          {
              If Missing
                  Missing = %Missing% %i%
              Else
                  Missing = %i%
  ;~ 	 		Msgbox Is: %A_Loopfield%`nShould be: %i%
          }
          i++
      }
      i++
  }
  ;~ msgbox Missing: %missing%
  If Missing =
    Missing = none
  vol =
  Volumelist =
  Filedelete, vol.txt
  Return, Missing
}



; ******************************************************************
; CMDret-AHK functions by corrupt
;
; CMDret_Stream
; version 0.03 beta
; Updated: Feb 19, 2007
;
; CMDret code modifications and/or contributions have been made by:
; Laszlo, shimanov, toralf, Wdb
; ******************************************************************
; Usage:
; CMDin - command to execute
; CMDname - type of output to process (Optional)
; WorkingDir - full path to working directory (Optional)
; ******************************************************************
; Known Issues:
; - If using dir be sure to specify a path (example: cmd /c dir c:\)
; or specify a working directory
; - Running 16 bit console applications may not produce output. Use
; a 32 bit application to start the 16 bit process to receive output
; ******************************************************************
; Additional requirements:
; - Your script must also contain a CMDret_Output function
;
; CMDret_Output(CMDout, CMDname="")
; Usage:
; CMDout - each line of output returned (1 line each time)
; CMDname - type of output to process (Optional)
; ******************************************************************
; Code Start
; ******************************************************************

CMDret_Stream(CMDin, CMDname="", WorkingDir=0)
{
  Global cmdretPID
  tcWrk := WorkingDir=0 ? "Int" : "Str"
  idltm := A_TickCount + 20
  LivePos = 1
  VarSetCapacity(CMDout, 1, 32)
  VarSetCapacity(sui,68, 0)
  VarSetCapacity(pi, 16, 0)
  VarSetCapacity(pa, 12, 0)
  Loop, 4 {
    DllCall("RtlFillMemory", UInt,&pa+A_Index-1, UInt,1, UChar,12 >> 8*A_Index-8)
    DllCall("RtlFillMemory", UInt,&pa+8+A_Index-1, UInt,1, UChar,1 >> 8*A_Index-8)
  }
  IF (DllCall("CreatePipe", "UInt*",hRead, "UInt*",hWrite, "UInt",&pa, "Int",0) <> 0) {
    Loop, 4
      DllCall("RtlFillMemory", UInt,&sui+A_Index-1, UInt,1, UChar,68 >> 8*A_Index-8)
    DllCall("GetStartupInfo", "UInt", &sui)
    Loop, 4 {
      DllCall("RtlFillMemory", UInt,&sui+44+A_Index-1, UInt,1, UChar,257 >> 8*A_Index-8)
      DllCall("RtlFillMemory", UInt,&sui+60+A_Index-1, UInt,1, UChar,hWrite >> 8*A_Index-8)
      DllCall("RtlFillMemory", UInt,&sui+64+A_Index-1, UInt,1, UChar,hWrite >> 8*A_Index-8)
      DllCall("RtlFillMemory", UInt,&sui+48+A_Index-1, UInt,1, UChar,0 >> 8*A_Index-8)
    }
    IF (DllCall("CreateProcess", Int,0, Str,CMDin, Int,0, Int,0, Int,1, "UInt",0, Int,0, tcWrk, WorkingDir, UInt,&sui, UInt,&pi) <> 0) {
      Loop, 4
        cmdretPID += *(&pi+8+A_Index-1) << 8*A_Index-8
      Loop {
        idltm2 := A_TickCount - idltm
        If (idltm2 < 15) {
          DllCall("Sleep", Int, 15)
          Continue
        }
        IF (DllCall("PeekNamedPipe", "uint", hRead, "uint", 0, "uint", 0, "uint", 0, "uint*", bSize, "uint", 0 ) <> 0 ) {
          Process, Exist, %cmdretPID%
          IF (ErrorLevel OR bSize > 0) {
            IF (bSize > 0) {
              VarSetCapacity(lpBuffer, bSize+1, 0)
              IF (DllCall("ReadFile", "UInt",hRead, "Str", lpBuffer, "Int",bSize, "UInt*",bRead, "Int",0) > 0) {
                IF (bRead > 0) {
                  IF (StrLen(lpBuffer) < bRead) {
                    VarSetCapacity(CMcpy, bRead, 32)
                    bRead2 = %bRead%
                    Loop {
                      DllCall("RtlZeroMemory", "UInt", &CMcpy, Int, bRead)
                      NULLptr := StrLen(lpBuffer)
                      cpsize := bread - NULLptr
                      DllCall("RtlMoveMemory", "UInt", &CMcpy, "UInt", (&lpBuffer + NULLptr + 2), "Int", (cpsize - 1))
                      DllCall("RtlZeroMemory", "UInt", (&lpBuffer + NULLptr), Int, cpsize)
                      DllCall("RtlMoveMemory", "UInt", (&lpBuffer + NULLptr), "UInt", &CMcpy, "Int", cpsize)
                      bRead2 --
                      IF (StrLen(lpBuffer) > bRead2)
                        break
                    }
                  }
              VarSetCapacity(lpBuffer, -1)
                  CMDout .= lpBuffer
                  bRead = 0
                }
              }
            }
          }
          ELSE
            break
        }
        ELSE
          break
        idltm := A_TickCount
        LiveFound := RegExMatch(CMDout, "m)^(.*)", LiveOut, LivePos)
        If (LiveFound)
          SetTimer, cmdretSTR, 5
      }
      cmdretPID=
      DllCall("CloseHandle", UInt, hWrite)
      DllCall("CloseHandle", UInt, hRead)
    }
  }
  StringTrimLeft, LiveRes, CMDout, %LivePos%
  If LiveRes <>
    Loop, Parse, LiveRes, `n
    {
      FileLine = %A_LoopField%
      StringTrimRight, FileLine, FileLine, 1
      CMDret_Output(FileLine, CMDname)
    }
  StringTrimLeft, CMDout, CMDout, 1
  cmdretPID = 0
  Return, CMDout
cmdretSTR:
SetTimer, cmdretSTR, Off
If (LivePosLast <> LiveFound) {
  FileLine = %LiveOut1%
  LivePos := LiveFound + StrLen(FileLine) + 1
  LivePosLast := LivePos
  CMDret_Output(FileLine, CMDname)
}
Return
}