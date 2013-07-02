VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisMacroStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub GlobalMacroStorage_Start()
    'On Error Resume Next
    Dim bar As CommandBar
    Dim ButLatex As CommandBarControl
    Dim ic As String
    
    Set bar = Application.CommandBars("Toolbox")
    Set ButLatex = bar.Controls.AddCustomButton(VGCore.UICategories.cdrCmdCategoryMacros, "LaTeX_CD.LaTeX_Macro.LatexEdit", temporary:=True)
    

    ButLatex.ToolTipText = "Latex-Objekt bearbeiten / erstellen"
    ButLatex.DescriptionText = "Latex-Objekt bearbeiten / erstellen"
    ButLatex.Caption = "LaTeX"
    'ID über Application.CommandBars(..).Controls(#).ID
    ic = SaveIcon
    
    
    'ButLatex.SetIcon2 ("guid://4a39f90c-ff2e-48d6-9283-16ccafd8dd89")
    ButLatex.SetCustomIcon (ic)
End Sub

Sub getIconStringOfButton()
    Dim dao As DataObject
    Set dao = New DataObject
    dao.SetText Application.CommandBars("Toolbox").Controls(20).ID
    
    dao.PutInClipboard
End Sub

Function SaveIcon() As String
    Dim encodedIcon As String
    Dim iconByte() As Byte
    Dim b64 As New Base64Class
    Dim sfilename As String
    Dim iFileNum As Integer
    Dim lWritePos As Long
    
    encodedIcon = "AAABAAEAEhIAAAAAAADYBQAAFgAAACgAAAASAAAAJAAAAAEACAAAAAAAsAEAAAAAAAAAAAAAAAAAAAAAAAAxMTEANjY2ADc3NwA+"
    encodedIcon = encodedIcon & "Pj4AQEBAAEtLSwBoaGgAbGxsAHBwcAB2dnYAeHh4AH19fQB/f38AgYGBAIKCggCEhIQAiIiIAI+PjwCSkpIAlZWVAJiYmACbm5sA"
    encodedIcon = encodedIcon & "nZ2dAJ+fnwChoaEAoqKiAKSkpACmpqYAqKioAKmpqQCtra0Arq6uAK+vrwCxsbEAs7OzALS0tADFxcUAxsbGAMzMzADPz88A0NDQ"
    encodedIcon = encodedIcon & "ANHR0QDV1dUA1tbWANjY2ADc3NwA3d3dAN7e3gDf398A4ODgAOHh4QDp6ekA7u7uAO/v7wDz8/MA9vb2APj4+AD5+fkA+vr6APz8"
    encodedIcon = encodedIcon & "/AD9/f0A/v7+AP///wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    encodedIcon = encodedIcon & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    encodedIcon = encodedIcon & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    encodedIcon = encodedIcon & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    encodedIcon = encodedIcon & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    encodedIcon = encodedIcon & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    encodedIcon = encodedIcon & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    encodedIcon = encodedIcon & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    encodedIcon = encodedIcon & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    encodedIcon = encodedIcon & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    encodedIcon = encodedIcon & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPj4+Pj4+Pj4+Pj4+Pj4+Pj4+AAA+Pj4+Pj4+Pj4+Pj4+Pj4+Pj4A"
    encodedIcon = encodedIcon & "AD4+Pj4+Pj4+Pj4+Pj4+Pj4+PgAAPj4+Pj4+Pj4+Pj4+Pj4+Pj4+AAA+Pj4+PjUUESIfDDE+Pj4+Pj4AAD4+Pj4+PiMOPj4tID4+"
    encodedIcon = encodedIcon & "Pj4+PgAAPjYhFSA0Iw4+NTcLGD0+HRcrAAA+Pj4APj4jCTIWPjUUPj4IGT4AAD4+PgI+PiMGJBI+PisqKAU+PgAAPj4+Aj4+Iw4+"
    encodedIcon = encodedIcon & "MzwmPiEELz4+AAA+Pj4CPj4hDT4+HSw+EQc+Pj4AACk+PgI+NRAaIh4VODADJi4+PgAAHDk+AT47Gz4+Pj4+BSc+Gjo+AAAmFCEP"
    encodedIcon = encodedIcon & "IRUlPj4+PhwKKj4REzYAAD4+Pj4+Pj4+Pj4+Pj4+Pj4+PgAAPj4+Pj4+Pj4+Pj4+Pj4+Pj4+AAA+Pj4+Pj4+Pj4+Pj4+Pj4+Pj4A"
    encodedIcon = encodedIcon & "AD4+Pj4+Pj4+Pj4+Pj4+Pj4+PgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    encodedIcon = encodedIcon & "AAAAAAAAAAAAAAAAAAAAAAAA"

    iconByte = b64.DecodeToByteArray(encodedIcon)
    
    sfilename = CorelDRAW.CorelScriptTools.GetTempFolder() & "latex.ico"
    lWritePos = 1
    iFileNum = FreeFile
    
    Open sfilename For Binary Access Write As #iFileNum
    Put #iFileNum, lWritePos, iconByte
    Close #iFileNum
    SaveIcon = sfilename
End Function
