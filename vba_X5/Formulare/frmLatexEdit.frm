VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLatexEdit 
   Caption         =   "Latex-Eingabe"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   OleObjectBlob   =   "frmLatexEdit.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmLatexEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VBA Macro for Corel Draw to easily import LaTeX-Output as vector Objects
'Macro is able to edit previously created LaTeX-Object while keeping position, scale, rotating and shearing
'Created by Jan Bender available at http://www.impulse-based.de/
'Modified for the use of Ghostscript and easy header-editing by Simon Klein (mail@simonklein.de)
'needs Ghostscript to be installed (path in registry or PATH-variable) and pdflatex in PATH (should already be with miktex and gpl ghostscript on windows)
'header stored in Registry

Option Explicit

Private Declare PtrSafe Function SetEnvironmentVariable Lib "kernel32" _
  Alias "SetEnvironmentVariableA" ( _
  ByVal lpName As String, _
  ByVal lpValue As String) As Long

Private Declare PtrSafe Function OpenProcess Lib "kernel32" _
(ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
ByVal dwProcessId As Long) As Long

Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const STATUS_PENDING = &H103&
Private Const PROCESS_QUERY_INFORMATION = &H400

Public defhead As String

Public Function ShellandWait(ExeFullPath As String, _
Optional TimeOutValue As Long = 0) As Boolean
    
    Dim lInst As Long
    Dim lStart As Long
    Dim lTimeToQuit As Long
    Dim sExeName As String
    Dim lProcessId As Long
    Dim lExitCode As Long
    Dim bPastMidnight As Boolean
    
    On Error GoTo ErrorHandler

    lStart = CLng(Timer)
    sExeName = ExeFullPath

    'Deal with timeout being reset at Midnight
    If TimeOutValue > 0 Then
        If lStart + TimeOutValue < 86400 Then
            lTimeToQuit = lStart + TimeOutValue
        Else
            lTimeToQuit = (lStart - 86400) + TimeOutValue
            bPastMidnight = True
        End If
    End If

    lInst = Shell(sExeName, vbMinimizedFocus)
    
    lProcessId = OpenProcess(PROCESS_QUERY_INFORMATION, False, lInst)

    Do
        Call GetExitCodeProcess(lProcessId, lExitCode)
        DoEvents
        If TimeOutValue And Timer > lTimeToQuit Then
            If bPastMidnight Then
                 If Timer < lStart Then Exit Do
            Else
                 Exit Do
            End If
    End If
    Loop While lExitCode = STATUS_PENDING
    
    ShellandWait = True
    Exit Function
   
ErrorHandler:
    ShellandWait = False
    Exit Function
End Function


Public Function ShellandWait2(ExeFullPath As String, _
Optional TimeOutValue As Long = 0) As Boolean
    
    Dim lInst As Long
    Dim lStart As Long
    Dim lTimeToQuit As Long
    Dim sExeName As String
    Dim lProcessId As Long
    Dim lExitCode As Long
    Dim bPastMidnight As Boolean
    
    On Error GoTo ErrorHandler

    lStart = CLng(Timer)
    sExeName = ExeFullPath

    'Deal with timeout being reset at Midnight
    If TimeOutValue > 0 Then
        If lStart + TimeOutValue < 86400 Then
            lTimeToQuit = lStart + TimeOutValue
        Else
            lTimeToQuit = (lStart - 86400) + TimeOutValue
            bPastMidnight = True
        End If
    End If

    lInst = Shell(sExeName, vbNormal)
    
    lProcessId = OpenProcess(PROCESS_QUERY_INFORMATION, False, lInst)

    Do
        Call GetExitCodeProcess(lProcessId, lExitCode)
        DoEvents
        If TimeOutValue And Timer > lTimeToQuit Then
            If bPastMidnight Then
                 If Timer < lStart Then Exit Do
            Else
                 Exit Do
            End If
    End If
    Loop While lExitCode = STATUS_PENDING
    
    ShellandWait2 = True
    Exit Function
   
ErrorHandler:
    ShellandWait2 = False
    Exit Function
End Function


Private Sub Cancel_Button_Click()
    Unload Me
End Sub


Private Sub CommandButton1_Click()
    frmLatexHeader.Show
End Sub

Private Sub Ok_Button_Click()
    Dim header As String
    header = GetSetting("ltxCrlEdt", "runtime", "header", defhead)
    
    Dim sOld As Shape
    Dim s As String
    Dim path As String
    Dim curpath As String
    Dim d11 As Double, d12 As Double, d21 As Double, d22 As Double
    Dim tx As Double, ty As Double
    Dim tex_template1 As String, tex_template2 As String, txt As String
    Dim MODE As Integer
    
    Dim impflt As ImportFilter
    Dim impopt As New StructImportOptions
    Dim s1 As Shape

'    MODE = 1 ' dvi (standard)
    MODE = 2 ' pdf
'    MODE = 3 ' dvi (ps -> ps2epsi)
    
    Set sOld = ActiveShape
    If Not (sOld Is Nothing) Then
        sOld.GetMatrix d11, d12, d21, d22, tx, ty
    End If
    path = Environ$("TEMP")
    curpath = CurDir
    s = TextBox1.Text
    path = path & "\"
    
    Open (path + "teximport.tex") For Output As #1
    Print #1, Replace(header, "%%ANCHOR%%", s)
    Close #1
    
    ChDrive path
    ChDir path
    
    On Error Resume Next
    'Kill "teximport.dvi"
    'Kill "teximport.ps"
    'Kill "teximport.pdf"
    'Kill "teximport.eps"
    
    txt = ActiveDocument.FilePath
    txt = Left(txt, Len(txt) - 1)
    If MODE = 2 Then
        If Not ShellandWait("pdflatex.exe --include-directory=""" + txt + """ teximport.tex") Then
            MsgBox "latex.exe nicht gefunden"
        End If
        If Dir("teximport.pdf", vbNormal) = "" Then
            MsgBox "pdflatex hat Datei teximport.pdf nicht erzeugt"
            Exit Sub
        End If
        On Error GoTo 0
        'Set environ Variable for miktex gs
        Dim miktex As String
        miktex = getMiktexPath
        
        Dim nulvar As Variant
        nulvar = SetEnvironmentVariable("MIKTEX_GS_LIB", miktex + "\..\..\ghostscript\base;" + miktex + "\..\..\fonts")
        
        'MsgBox (miktex)
        
        'Use Ghostscript to interprete
        If Not ShellandWait(miktex + "\mgs.exe -sDEVICE=pswrite -dNOCACHE -sOutputFile=teximport.ps -q -dbatch -dNOPAUSE teximport.pdf -c quit") Then
            If Not ShellandWait("gs.exe -sDEVICE=pswrite -dNOCACHE -sOutputFile=teximport.ps -q -dbatch -dNOPAUSE teximport.pdf -c quit") Then
                If Not ShellandWait(getGSPath + "gs.exe -sDEVICE=pswrite -dNOCACHE -sOutputFile=teximport.ps -q -dbatch -dNOPAUSE teximport.pdf -c quit") Then
                     MsgBox "Ghostscript nicht gefunden"
                End If
            End If
        End If
        
        If Dir("teximport.ps", vbNormal) = "" Then
            MsgBox "Datei teximport.ps nicht erzeugt"
            Exit Sub
        End If
        
        ChDrive curpath
        ChDir curpath
        
        impopt.MaintainLayers = True
        impopt.MODE = cdrImportFull
        Set impflt = ActiveLayer.ImportEx(path + "teximport.ps", cdrPSInterpreted, impopt)
        impflt.Finish
        Set s1 = ActiveShape
        Rem s1.SetPosition sOld.PositionX, sOld.PositionY
        If Not (sOld Is Nothing) Then
            s1.SetMatrix d11, d12, d21, d22, tx, ty
            sOld.Delete
        Else
            With ActiveWindow.ActiveView
                s1.SetPosition .OriginX, .OriginY
            End With
        End If
        
        s1.Name = "latex"
        s1.ObjectData("Comments") = s
        Unload Me
        
        Kill path + "teximport.*"
        Exit Sub
    End If

End Sub

Private Sub UserForm_Initialize()
    defhead = "\documentclass[pdflatex,12pt,a4paper]{report}" + vbCrLf
    defhead = defhead + "\usepackage{amsmath} " + vbCrLf
    defhead = defhead + "\usepackage{bbm} " + vbCrLf
    defhead = defhead + "\usepackage{commath} " + vbCrLf
    defhead = defhead + "\usepackage{textcomp} " + vbCrLf
    defhead = defhead + "\usepackage{sistyle} " + vbCrLf
    defhead = defhead + "\usepackage{bigstrut} " + vbCrLf
    defhead = defhead + "\usepackage{ae,aecompl} " + vbCrLf
    defhead = defhead + "\newcommand{\boldm}[1]{\mbox{\boldmath{$\mathrm {#1}$}}}" + vbCrLf
    defhead = defhead + "\def \b#1{\relax \ifmmode{\boldm {#1}\/}\else{$\bf {#1}\/$}\fi}" + vbCrLf
    defhead = defhead + "\begin{document} " + vbCrLf
    defhead = defhead + "\thispagestyle{empty}" + vbCrLf
    defhead = defhead + "%%ANCHOR%%" + vbCrLf
    defhead = defhead + "\end{document}" + vbCrLf

End Sub
