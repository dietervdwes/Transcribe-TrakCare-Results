;TrakCare Data entry Script
;Written by Dieter van der Westhuizen 08-05-2020
;dietervdwes@gmail.com
;dieter.vdwesthuizen@nhls.ac.za

#SingleInstance, force
;#NoTrayIcon
#NoEnv
#Persistent
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetKeyDelay, 150
SendMode, Input
SetTitleMatchMode, 2
SetDefaultMouseSpeed, 0
SetMouseDelay, 0
SetWinDelay, 500



;;;;;;;;;;;;;;;;;;;;;;;;;;;;    Add Buttons ;;;;;;;;;;;;;;;;;;
Gui, Add, button, x2 y2 w75 h20 ,Plasma_AA
Gui, add, button,x2 y22 w75 h20  ,Urine_AA
Gui, add, button,x2 y44 w75 h20  ,COV_2
;Gui, Add, button, x2 y42 w28 h20  ,EPR
;Gui, Add, button, x2 y64 w33 h20  ,FPSA
;Gui, Add, button, x2 y86 w45 h20  ,Verified
;Gui, Add, button, x2 y108 w57 h20  ,KeepOpen
;Gui, Add, button, w57 , VolVeri
;Gui, Add, button, x2 y130 w45 h20  ,More
Gui, Add, button, x2 y152 w32 h20 ,Close
Gui, Add, button, x45 y152 w20 h20 ,_i
;Gui, Add, Button, x6 y17 w100 h30 , Ok
;Gui, Add, Button, x116 y17 w100 h30 , Cancel
;;;;;;;;;;;;;;;;;;;;;;;;;;;;   Set Window Options   ;;;;;;;;;;;;;;
;Gui, +AlwaysOnTop
Gui, -sysmenu +AlwaysOnTop
Gui, Show, , Transcribe
WinGetPos,,,,TrayHeight,ahk_class Shell_TrayWnd,,,
height := A_ScreenHeight-270
width := A_ScreenWidth-85
Gui, Margin, 0, 0
;Gui, Add, Picture, x0 y0 w411 h485, picture.png
;Gui -Caption -Border
Gui, Show, x%width% y%height% w80

return

;;;;;;;;   This is to read the CSV called aa.csv, parse the data and send to TrakCare window, Results Entry.    

ButtonPlasma_AA:
sleep, 500
IfWinNotExist, Result Entry
        {
            MsgBox, Please open the Result Entry Window
            Return
        }
WinActivate, Result Entry
;^!U::
FileName := aa.csv
Loop, read, aa.csv, aa.txt
{
    LineNumber := A_Index
    Loop, parse, A_LoopReadLine, CSV
        
    if InStr(A_LoopField, "p.D")
    {
        EpisodeLine := LineNumber + 2
        FileReadLine, lEpisode_outr, aa.csv, EpisodeLine
        StringTrimRight, lEpisode_out, lEpisode_outr, 2
        StringTrimLeft, Episode_out, lEpisode_out, 1
        
        
        NorValineInternalStandardLine := LineNumber + 18
        AlphaaminobutyrateLine := LineNumber + 19
        AlanineLine := LineNumber + 20
        BetaaminoisobutyrateLine := LineNumber + 21
        AsparagineLine := LineNumber + 22
        AspartateLine := LineNumber + 23
        CystathionineLine := LineNumber + 24
        CystineCCdisulfideLine := LineNumber + 25
        GlutamateLine := LineNumber + 26
        GlutamineLine := LineNumber + 27
        GlycineLine := LineNumber + 28
        GlycylprolinedipeptideLine := LineNumber + 29
        HistidineLine := LineNumber + 30
        OHProlineLine := LineNumber + 31
        HydroxylysineLine := LineNumber + 32
        IsoleucineLine := LineNumber + 33
        LeucineLine := LineNumber + 34
        LysineLine := LineNumber + 35
        MethionineLine := LineNumber + 36
        OrnithineLine := LineNumber + 37
        PhenylalanineLine := LineNumber + 38
        ProlineLine := LineNumber + 39
        prolinehydroxyprolinedipeptideLine := LineNumber + 40
        SarcosineLine := LineNumber + 41
        SerineLine := LineNumber + 42
        ThioprolineLine := LineNumber + 43
        ThreonineLine := LineNumber + 44
        TryptophanLine := LineNumber + 45
        TyrosineLine := LineNumber + 46
        ValineLine := LineNumber + 47
        BetaAlanin_PLine := LineNumber + 48
        ArgininoSuccinateLine := LineNumber + 49
        HomocystineHcyHcyLine := LineNumber + 50
        ;FileAppend, Nor Valine Internal Standard is in line %NorValineInternalStandard%`n
        
        FileReadLine, NorValineInternalStandard_out, aa.csv, NorValineInternalStandardLine
        NorValineInternalStandardArray := StrSplit(NorValineInternalStandard_out, ",")
        NorValineInternalStandardResult := NorValineInternalStandardArray[4]
        FileReadLine, Alphaaminobutyrate_out, aa.csv, AlphaaminobutyrateLine
        AlphaaminobutyrateArray := StrSplit(Alphaaminobutyrate_out, ",")
        AlphaaminobutyrateResult := AlphaaminobutyrateArray[4]
        FileReadLine, Alanine_out, aa.csv, AlanineLine
        AlanineArray := StrSplit(Alanine_out, ",")
        AlanineResult := AlanineArray[4]
        FileReadLine, Betaaminoisobutyrate_out, aa.csv, BetaaminoisobutyrateLine
        BetaaminoisobutyrateArray := StrSplit(Betaaminoisobutyrate_out, ",")
        BetaaminoisobutyrateResult := BetaaminoisobutyrateArray[4]
        FileReadLine, Asparagine_out, aa.csv, AsparagineLine
        AsparagineArray := StrSplit(Asparagine_out, ",")
        AsparagineResult := AsparagineArray[4]
        FileReadLine, Aspartate_out, aa.csv, AspartateLine
        AspartateArray := StrSplit(Aspartate_out, ",")
        AspartateResult := AspartateArray[4]
        FileReadLine, Cystathionine_out, aa.csv, CystathionineLine
        CystathionineArray := StrSplit(Cystathionine_out, ",")
        CystathionineResult := CystathionineArray[4]
        FileReadLine, CystineCCdisulfide_out, aa.csv, CystineCCdisulfideLine
        CystineCCdisulfideArray := StrSplit(CystineCCdisulfide_out, ",")
        CystineCCdisulfideResult := CystineCCdisulfideArray[4]
        FileReadLine, Glutamate_out, aa.csv, GlutamateLine
        GlutamateArray := StrSplit(Glutamate_out, ",")
        GlutamateResult := GlutamateArray[4]
        FileReadLine, Glutamine_out, aa.csv, GlutamineLine
        GlutamineArray := StrSplit(Glutamine_out, ",")
        GlutamineResult := GlutamineArray[4]
        FileReadLine, Glycine_out, aa.csv, GlycineLine
        GlycineArray := StrSplit(Glycine_out, ",")
        GlycineResult := GlycineArray[4]
        FileReadLine, Glycylprolinedipeptide_out, aa.csv, GlycylprolinedipeptideLine
        GlycylprolinedipeptideArray := StrSplit(Glycylprolinedipeptide_out, ",")
        GlycylprolinedipeptideResult := GlycylprolinedipeptideArray[4]
        FileReadLine, Histidine_out, aa.csv, HistidineLine
        HistidineArray := StrSplit(Histidine_out, ",")
        HistidineResult := HistidineArray[4]
        FileReadLine, OHProline_out, aa.csv, OHProlineLine
        OHProlineArray := StrSplit(OHProline_out, ",")
        OHProlineResult := OHProlineArray[4]
        FileReadLine, Hydroxylysine_out, aa.csv, HydroxylysineLine
        HydroxylysineArray := StrSplit(Hydroxylysine_out, ",")
        HydroxylysineResult := HydroxylysineArray[4]
        FileReadLine, Isoleucine_out, aa.csv, IsoleucineLine
        IsoleucineArray := StrSplit(Isoleucine_out, ",")
        IsoleucineResult := IsoleucineArray[4]
        FileReadLine, Leucine_out, aa.csv, LeucineLine
        LeucineArray := StrSplit(Leucine_out, ",")
        LeucineResult := LeucineArray[4]
        FileReadLine, Lysine_out, aa.csv, LysineLine
        LysineArray := StrSplit(Lysine_out, ",")
        LysineResult := LysineArray[4]
        FileReadLine, Methionine_out, aa.csv, MethionineLine
        MethionineArray := StrSplit(Methionine_out, ",")
        MethionineResult := MethionineArray[4]
        FileReadLine, Ornithine_out, aa.csv, OrnithineLine
        OrnithineArray := StrSplit(Ornithine_out, ",")
        OrnithineResult := OrnithineArray[4]
        FileReadLine, Phenylalanine_out, aa.csv, PhenylalanineLine
        PhenylalanineArray := StrSplit(Phenylalanine_out, ",")
        PhenylalanineResult := PhenylalanineArray[4]
        FileReadLine, Proline_out, aa.csv, ProlineLine
        ProlineArray := StrSplit(Proline_out, ",")
        ProlineResult := ProlineArray[4]
        FileReadLine, prolinehydroxyprolinedipeptide_out, aa.csv, prolinehydroxyprolinedipeptideLine
        prolinehydroxyprolinedipeptideArray := StrSplit(prolinehydroxyprolinedipeptide_out, ",")
        prolinehydroxyprolinedipeptideResult := prolinehydroxyprolinedipeptideArray[4]
        FileReadLine, Sarcosine_out, aa.csv, SarcosineLine
        SarcosineArray := StrSplit(Sarcosine_out, ",")
        SarcosineResult := SarcosineArray[4]
        FileReadLine, Serine_out, aa.csv, SerineLine
        SerineArray := StrSplit(Serine_out, ",")
        SerineResult := SerineArray[4]
        FileReadLine, Thioproline_out, aa.csv, ThioprolineLine
        ThioprolineArray := StrSplit(Thioproline_out, ",")
        ThioprolineResult := ThioprolineArray[4]
        FileReadLine, Threonine_out, aa.csv, ThreonineLine
        ThreonineArray := StrSplit(Threonine_out, ",")
        ThreonineResult := ThreonineArray[4]
        FileReadLine, Tryptophan_out, aa.csv, TryptophanLine
        TryptophanArray := StrSplit(Tryptophan_out, ",")
        TryptophanResult := TryptophanArray[4]
        FileReadLine, Tyrosine_out, aa.csv, TyrosineLine
        TyrosineArray := StrSplit(Tyrosine_out, ",")
        TyrosineResult := TyrosineArray[4]
        FileReadLine, Valine_out, aa.csv, ValineLine
        ValineArray := StrSplit(Valine_out, ",")
        ValineResult := ValineArray[4]
        FileReadLine, BetaAlanin_P_out, aa.csv, BetaAlanin_PLine
        BetaAlanin_PArray := StrSplit(BetaAlanin_P_out, ",")
        BetaAlanin_PResult := BetaAlanin_PArray[4]
        FileReadLine, ArgininoSuccinate_out, aa.csv, ArgininoSuccinateLine
        ArgininoSuccinateArray := StrSplit(ArgininoSuccinate_out, ",")
        ArgininoSuccinateResult := ArgininoSuccinateArray[4]
        FileReadLine, HomocystineHcyHcy_out, aa.csv, HomocystineHcyHcyLine
        HomocystineHcyHcyArray := StrSplit(HomocystineHcyHcy_out, ",")
        HomocystineHcyHcyResult := HomocystineHcyHcyArray[4]
        
        
        IfWinActive, Result Entry - Single
            WinClose, Result Entry - Single
        IfWinNotActive, Result Entry
        {
            MsgBox, Please open the Result Entry Window
            Return
        }
        
        ;MsgBox, Now going to transcribe results for %Episode_out%
        
        WinActivate, Result Entry
        WinWaitActive, Result Entry
        ;Run, Notepad
        ;WinWaitActive, Notepad
        sleep, 500
        send, {AltDown}l{AltUp}
        sleep, 300
        ;MouseClick, Left, 105, 112, 2, 100
        ;MouseClick, Left, 105, 500, 2, 100
        sleep, 500
        send, %Episode_out%
        sleep, 300
        send, {Tab}
        sleep, 300
        send, {Tab}
        sleep, 300
        send, AAQ
        sleep, 200
        send, {Enter}
        sleep, 1000
        MouseClick, Left, 150, 300
        sleep, 1000
        send, {AltDown}e{AltUp}
        WinWaitActive, Result Entry - Single
        
        sleep, 3000
        send, %AlanineResult%
        Loop, 3
        {
            send, {Tab}
            sleep, 300
        }
        send, %AlphaaminobutyrateResult%
        Loop, 6
        {
            send, {Tab}
            sleep, 300
        }
        send, %AsparagineResult%
        sleep, 300
        send, {Tab}
        sleep,300
        send, %AspartateResult%
        sleep,300
        Loop, 4
        {
            send, {Tab}
            sleep, 300
        }
        send, %CystineCCdisulfideResult%
        sleep, 100
        send, {TAB}
        sleep, 300
        send, %GlutamateResult%
        sleep, 100
        send, {TAB}
        sleep, 300
        send, %GlutamineResult%
        sleep, 100
        send, {TAB}
        sleep, 300
        send, %GlycineResult%
        sleep, 100
        send, {TAB}
        sleep, 300
        send, %HistidineResult%
        sleep, 100
        Loop, 5
        {
            send, {Tab}
            sleep, 300
        }
        send, %IsoleucineResult%
        sleep, 100
        send, {TAB}
        sleep, 300
        send, %LeucineResult%
        sleep, 100
        send, {TAB}
        sleep, 300
        send, %LysineResult%
        sleep, 100
        send, {TAB}
        sleep, 300
        send, %MethionineResult%
        sleep, 100
        send, {TAB}
        sleep, 300
        send, %OrnithineResult%
        sleep, 100
        send, {TAB}
        sleep, 300
        send, %PhenylalanineResult%
        sleep, 100
        send, {TAB}
        sleep, 300
        send, %ProlineResult%
        sleep, 100
        Loop, 2
        {
            send, {Tab}
            sleep, 300
        }
        send, %SerineResult%
        sleep, 100
        send, {TAB}
        sleep, 300
        send, %ThreonineResult%
        sleep, 100
        send, {TAB}
        sleep, 300
        send, %TryptophanResult%
        sleep, 100
        send, {TAB}
        sleep, 300
        send, %TyrosineResult%
        sleep, 100
        send, {TAB}
        sleep, 300
        send, %ValineResult%
        sleep, 100
        send, {TAB}
        sleep, 300
        send, {F6}
        WinWaitActive, Comments
        sleep, 300
        send, \aaq{Space}
        sleep, 500
        WinClose, Comments
        sleep, 300
        send, {Tab}
        sleep, 200
        send, y
        sleep, 300
        send, {Tab}
        sleep, 500
        send, {AltDown}u{AltUp}
        sleep, 1500
        */
        send, {AltDown}c{AltUp}
        sleep, 1000
        send, {AltDown}l{AltUp}
        sleep, 800
     }
}
return

ButtonUrine_AA:

IfWinNotExist, Result Entry
        {
            MsgBox, Please open the Result Entry Window, enter the WorkSheet as: SGCAAU and the number underneath and click "Find". 
            Return
        }


#IfWinActive, Result Entry - Single - 
;MsgBox, Please make sure that Result Entry window has already been opened and that the WorkSheet entered: SGCAAU and the number has been entered underneath it.`nClick OK when done, or "ESC" if this hasn't been done yet.

MsgBox, 4, , Please make sure that Result Entry window has already been opened and that the WorkSheet entered: SGCAAU and the number has been entered underneath it. Then Click on the first episode's result and Click Enter. `nAre you sure you want to continue?
IfMsgBox No
    return
else IfMsgBox Yes 
    InputBox, qty, Episodes, Enter number (quantity) of Episodes on worksheet, excluding QC:
sleep, 100

WinActivate, Result Entry
WinWaitActive, Result Entry
sleep, 2000

loop, %qty% ; how many times the process should be repeated
{
	MouseClick, left, 1011, 696, 1 ;Click scroll down button
	sleep, 500
	MouseClick, left, 1011, 696, 1 ;Click scroll down button
	sleep, 500
	MouseClick, left, 256, 644, 1 ;Click "Interpretation" box
	sleep, 500
	send, AMINO ;Types "AMINO" into above box
	sleep, 500
	send, {TAB}
	send, {F6}
	WinWaitActive, Comments
	sleep, 200
	send, \amino{Space} ;Types the amino comment in the F6 comment field.
	sleep, 500
	send, {AltDown}{F4}{AltUp} ; closes the window
	sleep, 500
	send, {Tab}
	sleep, 200
	send, Y
	sleep, 200
	send, {Tab}
	sleep, 500
	Send, {AltDown}a{AltUp} ; Authorizes
	sleep, 2500 ; waits for TrakCare to load Creatinine
	
	send, {AltDown}>{AltUp}
	sleep, 200
	send, {Enter}
	sleep, 7000 ;waits for Trak to load next Urine amino acid.
}
#IfWinActive
return

ButtonCOV_2:

FileName := COV2.csv
Loop, read, COV2.csv, ;output_list.txt ; output_list.txt is the file to write to, if necessary
{   ;This part of code reads and parses the CSV / txt file and stores results of the read file in an array.
    LineNumber := A_Index
    Loop, parse, A_LoopReadLine, CSV
        EpisodeLine := LineNumber
        FileReadLine, Line_out, COV2.csv, EpisodeLine
        LineArray := StrSplit(Line_out, ",")
        Episode_out := LineArray[1]
        Result := LineArray[2] 
       
        WinActivate, Result Entry
        WinWaitActive, Result Entry
        sleep, 800
        ;MouseClick, Left, 105, 112, 2, 100
        ;MouseClick, Left, 105, 500, 2, 100
        send, %Episode_out%
        sleep, 300
        send, {Tab}
        sleep, 300
        send, {Tab}
        sleep, 300
        send, V560
        sleep, 350
        send, {Enter}
        sleep, 500
        IfWinActive, Search Unsuccessful
        {
            send, {Enter}
            sleep, 500
            send, {AltDown}l{AltUp}
            FileAppend, %Episode_out%`,%Result%`,%A_Now%`,Search Unsuccessful`n, COV2.txt
            continue
        }
        else
        MouseClick, Left, 150, 300
        sleep, 500
        send, {AltDown}e{AltUp}
        sleep, 2000
        WinWaitActive, Result Entry - Single
        sleep, 1000 ;                                 -----------Increase time if Result Window takes too long to load
        sleep, 500
        PixelGetColor, color, 272, 265
        While !(color = 0xFFFF00)
        {
        PixelGetColor, color, 272, 265
        ToolTip, Color is %color%
        sleep, 300
        Tooltip
        }
        sleep, 100
        WinActivate, Result Entry - Single
        sleep, 100
        send, %Result%
        sleep, 200
        send, {Tab}
        sleep, 300
        send, %Result%
        sleep, 300
        send, {Tab}
        sleep, 300
        send, REF
        send, {Tab}
        sleep, 500
        send, {AltDown}a{AltUp}
        WinWaitClose, Result Entry - Single
        sleep, 1000
        FileAppend, %Episode_out%`,%Result%`,%A_Now%`n, COV2.txt
        ;sleep, 3000
        WinActivate, Result Entry
        ;WinWaitActive, Result Entry
        sleep, 1000
        send, {AltDown}l{AltUp}
        sleep, 300
}
        
Escape::Reload
;ExitApp
Return

^!r::Reload  ; Assign Ctrl-Alt-R as a hotkey to restart the script.

ButtonClose:
ExitApp

Button_i:
Run, http://github.com/dietervdwes/chemhelp


    

