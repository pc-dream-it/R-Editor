;@A.DEGARDIN 2019, 
#SingleInstance Force
#NoEnv
SetBatchLines -1

CurrentVersion := "1.1.01"
Update(CurrentVersion)


SetWorkingDir, A_ScriptDir ;  

Menu, tray, icon, %A_WinDir%\system32\shell32.dll, 262
Menu, FileMenu, Add, &Execute File, Execute_File
Menu, FileMenu,Icon,&Execute File, %A_WinDir%\system32\SHELL32.dll, 262 
Menu, FileMenu, Add, &New (Ctrl+N), FileNew
Menu, FileMenu, Add, &Open (Ctrl+O), FileOpen
Menu, FileMenu, Add, &Save (Ctrl+S), FileSave
Menu, FileMenu, Add, Save &As, FileSaveAs
Menu, FileMenu, Add  ; Separator line.
Menu, FileMenu, Add, &Exit, FileExit
Menu, HelpMenu, Add, &About R-Editor, HelpAbout
Menu, HelpMenu, Add, &Version, Version
;Menu, HelpMenu, Add, &About R_Editor, R_editor
;Menu, Execute, Add, &Execute, Execute_R
;Menu, Schedule, Add, &Schedule Daily, ScheduleDaily
;************************INSERT FUNCTIONS******************************************
Menu,Graphs,Add,Pareto Chart,          Pareto
Menu,Graphs,Add,Run Chart,          RunChart
Menu,Graphs,Add,Histogram Capability,          HistoCapa
Menu,Projects,Add,C&E Diagram,          Fishbone
Menu,Statistics,Add,Gage R and R,          GageR
Menu,Statistics,Add,Capability Analysis,          Capability

Menu, MyMenuBar, Add, &File, :FileMenu
Menu, MyMenuBar,Icon,&File, %A_WinDir%\system32\SHELL32.dll, 7
Menu, MyMenuBar, Add
Menu, MyMenuBar, Add

Menu, MyMenuBar, Add, &Graphs, :Graphs
Menu, MyMenuBar,Icon,&Graphs, %A_WinDir%\system32\SHELL32.dll, 167 
Menu, MyMenuBar, Add, &Projects, :Projects
Menu, MyMenuBar,Icon,&Projects, %A_WinDir%\system32\SHELL32.dll, 166 
Menu, MyMenuBar, Add, &Statistics, :Statistics
Menu, MyMenuBar,Icon,&Statistics, %A_WinDir%\system32\SHELL32.dll, 270 
Menu, MyMenuBar, Add
Menu, MyMenuBar, Add
Menu, MyMenuBar, Add, &Run, Execute_R
Menu, MyMenuBar,Icon,&Run, %A_WinDir%\system32\SHELL32.dll, 138 
Menu, MyMenuBar, Add, &Stop, Stop_exe
Menu, MyMenuBar,Icon,&Stop, %A_WinDir%\system32\SHELL32.dll, 216 
Menu, MyMenuBar, Add
Menu, MyMenuBar, Add

Menu, MyMenuBar, Add, &Help, :HelpMenu
Menu, MyMenuBar,Icon,&Help, %A_WinDir%\system32\SHELL32.dll, 155 

;Menu, MyMenuBar, Add, &Open Log, Log
;Menu, MyMenuBar,Icon,&Open Log, %A_WinDir%\system32\SHELL32.dll, 166 
;Menu, MyMenuBar, Add, &Schedule, :Schedule
;Menu, MyMenuBar,Icon,&Schedule, %A_WinDir%\system32\SHELL32.dll, 240 

; Attach the menu bar to the window:
Gui, Menu, MyMenuBar

; Create the main Edit control and display the window:
Gui, +Resize  ; Make the window resizable.

Gui Add, Edit, vMainEdit x13 y39 w418 h413 +Multi
Gui Add, Edit, vRes x438 y39 w427 h413 +Multi
Gui Add, Text, vFilenametext x13 y6 w525 h23 +0x200, Untitled

Gui Show, w873 h466, R_Editor for Six Sigma
Return

CurrentFileName =  ; Indicate that there is no current file.



HistoCapa:
Store:=ClipboardAll  ;****Store clipboard ****
Clipboard=
(  Join`r`n
#https://cran.r-project.org/web/packages/qcc/vignettes/qcc_a_quick_tour.html
library(qcc)
data(pistonrings)
diameter = with(pistonrings, qcc.groups(diameter, sample))
q1 = qcc(diameter[1:25,], type="xbar", nsigmas=3, plot=FALSE)

process.capability(q1, spec.limits=c(73.95,74.05))
)
Gosub Paste_and_Restore_Stored_Clipboard
return


RunChart:
Store:=ClipboardAll  ;****Store clipboard ****
Clipboard=
(  Join`r`n
#https://cran.r-project.org/web/packages/qcc/vignettes/qcc_a_quick_tour.html
library(qcc)
data(pistonrings)
diameter = with(pistonrings, qcc.groups(diameter, sample))
q1 = qcc(diameter[1:25,], type="xbar", newdata=diameter[26:40,],
         confidence.level=0.99)
plot(q1, restore.par = FALSE)
#add warning limits
#warn.limits = limits.xbar(q1$center, q1$std.dev, q1$sizes, 2)
#abline(h = warn.limits, lty = 3, col = "chocolate")
)
Gosub Paste_and_Restore_Stored_Clipboard
return


Pareto:
Store:=ClipboardAll  ;****Store clipboard ****
Clipboard=
(  Join`r`n
library(qcc)
#pareto chart of defects:
defect = c(80, 27, 66, 94, 33)
names(defect) = c("price code", "schedule date", "supplier code", "contact num.", "part num.")
pareto.chart(defect, ylab = "Error frequency")
)
Gosub Paste_and_Restore_Stored_Clipboard
return

Capability:
Store:=ClipboardAll  ;****Store clipboard ****
Clipboard=
(  Join`r`n
#https://datascienceplus.com/six-sigma-dmaic-series-in-r-part-2/

library(SixSigma)
# Create data for Process capability analysis
x<-c(755.81, 750.54, 751.05, 749.52, 749.21, 748.38,
              748.11, 753.07, 749.56, 750.08, 747.16, 747.53,
              749.22, 746.76, 747.64, 750.46, 749.27, 750.33,
              750.26, 751.29)

ss.ca.cp(x,740, 760)
ss.ca.cpk(x,740, 760)

# perform Capability Study
ss.study.ca(x, LSL = 740, USL = 760,
            Target = 750, alpha = 0.5,
            f.su = "Food Sample Example")
)
Gosub Paste_and_Restore_Stored_Clipboard
return


Fishbone:
Store:=ClipboardAll  ;****Store clipboard ****
Clipboard=
(  Join`r`n
#Cause and Effect fishbones
library(qcc)
cause.and.effect(cause = list(Measurements = c("Micrometers", "Microscopes", "Inspectors"),
                              Materials    = c("Alloys", "Lubricants", "Suppliers"),
                              Personnel    = c("Shifts", "Supervisors", "Training", "Operators"),
                              Environment  = c("Condensation", "Moisture"),
                              Methods      = c("Brake", "Engager", "Angle"),
                              Machines     = c("Speed", "Lathes", "Bits", "Sockets")),
                 effect = "Surface Flaws")
)
Gosub Paste_and_Restore_Stored_Clipboard
return


GageR:
Store:=ClipboardAll  ;****Store clipboard ****
Clipboard=
(  Join`r`n
#https://datascienceplus.com/six-sigma-dmaic-series-in-r-part-2/

# gage R & R study 

# Create gage R and R data for 3 piston rings , 2 operators and each operator 3 measurements per piston
Operator<- factor(rep(1:2, each = 9))
Pistonring<- factor(rep(rep(1:3, each = 3), 2))
run<- factor(rep(1:3, 6))
diameter<-c(1.4727, 1.4206, 1.4754, 1.5083, 1.5739,
            1.4341, 1.5517, 1.5483, 1.4614, 1.3337,
            1.6078, 1.4767, 1.4066, 1.5951, 1.8419,
            1.7087, 1.8259, 1.5444)
pistondata<-data.frame(Operator,Pistonring,run,diameter)
#View(pistondata)


#Load package
library("SixSigma")

#Perform gage R & R
my.rr <- ss.rr(var = diameter, part = Pistonring,      appr = Operator,       data = pistondata,  main = "Six Sigma Gage R&R Measure",   sub = "Piston ring MSA")
)
Gosub Paste_and_Restore_Stored_Clipboard
return


FileNew:
    GuiControl,, MainEdit  ; Clear the Edit control.
    Gui, Show,, R_Editor for Six Sigma
    CurrentFileName =  ; Indicate that there is no current file.
    GuiControl,, MainEdit, 
    GuiControl,, Filenametext, Untitled
    
    
return

Stop_exe:

WinClose, ahk_class ConsoleWindowClass

return



FileOpen:
Gui +OwnDialogs  ; Force the user to dismiss the FileSelectFile dialog before returning to the main window.
FileSelectFile, SelectedFileName, 3,, Open File ;, Text Documents (*.txt)
if SelectedFileName =  ; No file selected.
    return
Gosub FileRead
return

FileRead:  ; Caller has set the variable SelectedFileName for us.
FileRead, MainEdit, %SelectedFileName%  ; Read the file's contents into the variable.
if ErrorLevel
{
    MsgBox Could not open "%SelectedFileName%".
    return
}
GuiControl,, MainEdit, %MainEdit%  ; Put the text into the control.
CurrentFileName = %SelectedFileName%
GuiControl,, Filenametext, %SelectedFileName%

;Gui, Show,, %CurrentFileName%   ; Show file name in title bar.
return


^N::
Goto FileNew
return


^O::
Goto FileOpen
return


^S::
Goto FileSave
return

FileSave:
if CurrentFileName =   ; No filename selected yet, so do Save-As instead.
    Goto FileSaveAs
Gosub SaveCurrentFile
return

FileSaveAs:
Gui +OwnDialogs  ; Force the user to dismiss the FileSelectFile dialog before returning to the main window.
FileSelectFile, SelectedFileName, S16,, Save File ;, Text Documents (*.txt)
;FileSelectFile, SelectedFileName, 1,, Save File, AutoHotkey Scripts (*.ahk)

if SelectedFileName =  ; No file selected.
    return
CurrentFileName = %SelectedFileName%
GuiControl,, Filenametext, %SelectedFileName%
Gosub SaveCurrentFile
return

SaveCurrentFile:  ; Caller has ensured that CurrentFileName is not blank.
IfExist %CurrentFileName%
{
    FileDelete %CurrentFileName%
    if ErrorLevel
    {
        MsgBox The attempt to overwrite "%CurrentFileName%" failed.
        return
    }
}
GuiControlGet, MainEdit  ; Retrieve the contents of the Edit control.
FileAppend, %MainEdit%, %CurrentFileName% ;, AutoHotkey Scripts (*.ahk) ; Save the contents to the file.
; Upon success, Show file name in title bar (in case we were called by FileSaveAs):
;Gui, Show,, %CurrentFileName%

return

HelpAbout:
    run, https://github.com/pc-dream-it/R-Editor
return



Execute_R:

WinClose, ahk_class ConsoleWindowClass

;if CurrentFileName =   ; No filename selected yet, so do Save-As instead.
;    Goto FileSaveAs
;Gosub SaveCurrentFile

    launchR(CurrentFileName)
    WinWait, ahk_exe cmd.exe
    WinMove, ahk_exe cmd.exe ,,(A_ScreenWidth)-(600) ,(A_ScreenHeight)-(400) , 600, 400
return


Execute_File:

WinClose, ahk_class ConsoleWindowClass


Gui +OwnDialogs  ; Force the user to dismiss the FileSelectFile dialog before returning to the main window.
FileSelectFile, SelectedFileName, 3, , Open a file
;FileSelectFile, SelectedFileName, S16,, open File )

if SelectedFileName =  ; No file selected.
    return
CurrentFileName = %SelectedFileName%
GuiControl,, Filenametext, %SelectedFileName%

SelectedFileName := StrReplace(SelectedFileName, "\", "/")

GuiControl,, Edit1, source('%SelectedFileName%')


intro=
(  Join`r`n
pause = function()
{
    if (interactive())
    {
       #Sys.sleep(600)
       
    }
    else
    {
       invisible(readline(prompt = "STOP execution with Menu..."))
       Sys.sleep(600)

    }
}

X11()
interactive()
)
FoundPos := InStr(fileR, "\" , false,  0,1)
 Rpath := SubStr(fileR, 1, FoundPos)"scritptest.Rout"   

;FileRead, MainEdit, %fileR%

;take current  code    from code area
GuiControlGet, MainEdit ;Edit1 
;    GuiControlGet, sCurrentCode,, Edit1
Sleep, 200

;newcontent =''
newcontent = %intro% `r`n %MainEdit% `r`n pause()
;MsgBox, %newcontent%


tempfilepath := SubStr(fileR, 1, FoundPos)"tempfile.R"  

FileDelete, %tempfilepath%
FileAppend, %newcontent%, %tempfilepath%

;MsgBox, %newcontent%


;MsgBox, %Rpath%
        DelTask=R.exe CMD BATCH --vanilla --interactive "%tempfilepath%" "%Rpath%" 
        

        Run,  %DelTask%,, Min
        
        


/*
FoundPos := InStr(CurrentFileName, "\" , false,  0,1)
 Rpath := SubStr(CurrentFileName, 1, FoundPos)"scritptest.Rout"   
;MsgBox, %Rpath%
          
        DelTask=R.exe CMD BATCH --vanilla --interactive "%CurrentFileName%" "%Rpath%"  
        Run,  %DelTask%,, Min
*/
WinWaitActive, ahk_exe Rterm.exe 
if ErrorLevel
{
    Sleep, 200
    return
}

WinWaitNotActive, ahk_exe Rterm.exe 
if ErrorLevel
{
    Sleep, 200
    return
}
;keep current text output    
    GuiControlGet, sCurrentText,, Res
    
;append result from output
 FileRead, OutputVar, %Rpath%
    
    ;skip intro R
FoundPosIntro := InStr(OutputVar, "to quit R" , false,  0,1) 
;MsgBox, %FoundPosIntro%
 OutputVar := SubStr(OutputVar, FoundPosIntro+11)  
    
GuiControl,, Res, %sCurrentText% %OutputVar% 
    Sleep, 200
ControlSend, Edit2, ^{End}, R_Editor for Six Sigma

    WinWait, ahk_exe cmd.exe
    WinMove, ahk_exe cmd.exe ,,(A_ScreenWidth)-(600) ,(A_ScreenHeight)-(400) , 600, 400
return


;Log:
    ;run, %CurrentFileName%.log
;    FoundPos := InStr(CurrentFileName, "\" , false,  0,1)
;    Rpath := SubStr(CurrentFileName, 1, FoundPos)"scritptest.Rout"   
;MsgBox, %Rpath%
          
;        Run, notepad  "%Rpath%" 

;return

/*
ScheduleDaily:

    GuiControlGet , mFile,, Filenametext

    If InStr(mFile, " ")
    {
        MsgBox, Path:%mFile% `n Phantomjs does not work with file paths that have spaces in them. Please Save it in other Path without spaces
    }
    else
    {

    Gui 2:Add, Text, x19 y9 w139 h23 +0x200, Add Daily Task:
    Gui 2:Add, Text, x16 y43 w120 h23 +0x200, Task Name:
    Gui 2:Add, Edit, vthistaskname x16 y71 w399 h21, My task
    Gui 2:Add, Text, x16 y99 w120 h23 +0x200, Batch File (.bat) Path:
    Gui 2:Add, Edit, vfile x16 y129 w335 h21 
    Gui 2:Add, Button, x354 y129 w70 h21 gBROWSE, BROWSE
    Gui 2:Add, DateTime, vtime x15 y156 w120 h23   1, HH:mm
    Gui 2:Add, Button, x15 y233 w80 h23 gAddTaskDaily, ADD TASK
    Gui 2:Add, Button, x117 y232 w85 h23 gcheckTasks, CHECK TASKS
    Gui 2:Add, Button, x261 y232 w153 h23 gOpenScheduler, OPEN TASK SCHEDULER
    ;MsgBox, %mFile%
    GuiControl,2:, file, %mFile%.bat  
    
    Gui 2:Show, w426 h270, Tasks Scheduler Windows
    }

Return

2AddTaskDaily:  
2OpenScheduler:  
2GuiClose:
2GuiEscape:
Gui, 1:-Disabled  ; Re-enable the main window (must be done prior to the next step).
Gui, 2: Destroy  ; Destroy the about box.
return

*/

BROWSE()
{
FileSelectFile, SelectedFile, 3, , Open a .bat file 
GuiControl,2:, file, %SelectedFile%
}


AddTaskDaily:
    
    GuiControlGet , sFile,, file
    GuiControlGet , staskname,, thistaskname
    GuiControlGet , stime,, time
    FormatTime, stime, %stime%, HH:mm
    sfrequency=DAILY
    ;MsgBox,  SchTasks /Create /SC %sfrequency% /TN "%staskname%" /TR "%sfile%" /ST %stime%
    SheduleTask(staskname, sFile, stime, sfrequency)
return


;**************************SCHEDULER WINDOWS*****************

SheduleTask(mytaskname, file, time, frequency)
{

        DelTask=SCHTASKS.exe /Create /SC %frequency% /TN "%mytaskname%" /TR "%file%" /ST %time%
        Run,  %DelTask%

}

checkTasks:
    MsgBox,% GetTaskInfos()
return

GetTaskInfos()
{
    objService := ComObjCreate("Schedule.Service")
    objService.Connect()
    rootFolder := objService.GetFolder("\")
    taskCollection := rootFolder.GetTasks(0)
    numberOfTasks := taskCollection.Count
    ; ?RegistrationInfo.Author
    For registeredTask, state in taskCollection
    {
        if (registeredTask.state == 0)
            state:= "Unknown"
        else if (registeredTask.state == 1)
            state:= "Disabled"
        else if (registeredTask.state == 2)
            state:= "Queued"
        else if (registeredTask.state == 3)
            state:= "Ready"
        else if (registeredTask.state == 4)
            state:= "Running"
        tasklist .= registeredTask.Name "=" state "`n" ; "=" registeredTask.state "`n"
    }
    return tasklist
}
;**************************************************

OpenScheduler:
    run, control schedtasks
return


GuiSize:
if ErrorLevel = 1  ; The window has been minimized.  No action needed.
    return
; Otherwise, the window has been resized or maximized. Resize the Edit control to match.
;NewWidth := A_GuiWidth - 330
;NewHeight := A_GuiHeight - 50
;GuiControl, Move, MainEdit, W%NewWidth% H%NewHeight%
;NewWidth2 := A_GuiWidth - 30
;GuiControl, Move, Res, W%NewWidth2% H%NewHeight%
return


FileExit:     ; User chose "Exit" from the File menu.
GuiClose:  ; User closed the window.
    ExitApp


Esc::ExitApp

;************R Editor FUNCTIONS****************


launchR(fileR)
{

intro=
(  Join`r`n
pause = function()
{
    if (interactive())
    {
       #Sys.sleep(600)
       
    }
    else
    {
       invisible(readline(prompt = "STOP execution with Menu..."))
       Sys.sleep(600)

    }
}

X11()
interactive()
)
FoundPos := InStr(fileR, "\" , false,  0,1)
 Rpath := SubStr(fileR, 1, FoundPos)"scritptest.Rout"   

;FileRead, MainEdit, %fileR%

;take current  code    from code area
GuiControlGet, MainEdit ;Edit1 
;    GuiControlGet, sCurrentCode,, Edit1
Sleep, 200

;newcontent =''
newcontent = %intro% `r`n %MainEdit% `r`n pause()
;MsgBox, %newcontent%


tempfilepath := SubStr(fileR, 1, FoundPos)"tempfile.R"  

FileDelete, %tempfilepath%
FileAppend, %newcontent%, %tempfilepath%

;MsgBox, %newcontent%
filedelete, %Rpath%
sleep, 100

;MsgBox, %Rpath%
        DelTask=R.exe CMD BATCH --vanilla --interactive "%tempfilepath%" "%Rpath%" 
        

        Run,  %DelTask%,, Min
 
 
;keep current text output    
GuiControlGet, sCurrentText,, Res

WinWaitActive, ahk_exe Rterm.exe 
if ErrorLevel
{
    Sleep, 200
    return
}

WinWaitNotActive, ahk_exe Rterm.exe 
if ErrorLevel
{
    Sleep, 200
    return
}

;append result from output
FileRead, OutputVar, %Rpath%
    
FoundPosIntro := InStr(OutputVar, "to quit R" , false,  0,1) 
;MsgBox, %FoundPosIntro%
 OutputVar := SubStr(OutputVar, FoundPosIntro+11)  

 WinActivate, ahk_class AutoHotkeyGUI


GuiControl,, Res, %sCurrentText% %OutputVar% ^{End}
    Sleep, 200
 
; ControlSend, Edit2, ^{End}
        


}



;*************************************************************************

;***********Restore clipboard******************* 
Paste_and_Restore_Stored_Clipboard:  ;~ MsgBox % clipboard

ControlSend, Edit1, `r`n , R_Editor for Six Sigma
Sleep, 50
Send, ^v ;Depending on your OS and Admin level- you might want to check this
Sleep, 50
Clipboard:=Store  ;Restore clipboard to original contents
return

;****************************************

Update(CurrentVersion)
{
;https://github.com/pc-dream-it/R-Editor

    whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    whr.Open("GET", "https://raw.githubusercontent.com/pc-dream-it/R-Editor/master/version.txt", true)
    whr.Send()
    ; Using 'true' above and the call below allows the script to remain responsive.
    whr.WaitForResponse()
    version := whr.ResponseText
   
    if (version=CurrentVersion)
    {
    ;MsgBox,, R Editor version %version%, You have the last version
    }
    else
    {
    MsgBox,, R Editor %CurrentVersion%, A newer version is available %version% at https://github.com/pc-dream-it/R-Editor
    ;run, https://github.com/adegard/TagIE.ahk
    }
}


;******************************
Version:

    whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    whr.Open("GET", "https://raw.githubusercontent.com/pc-dream-it/R-Editor/master/version.txt", true)
    whr.Send()
    ; Using 'true' above and the call below allows the script to remain responsive.
    whr.WaitForResponse()
    version := whr.ResponseText
    MsgBox,, R Editor version %version%, You have the last version
return

