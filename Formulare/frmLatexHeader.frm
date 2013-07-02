VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLatexHeader 
   Caption         =   "LaTeX Header bearbeiten"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6825
   OleObjectBlob   =   "frmLatexHeader.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmLatexHeader"
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

Private Sub Cancel_Button_Click()
    Unload Me
End Sub

Private Sub CommandButton1_Click()
    Me.TextBox1.Value = frmLatexEdit.defhead
End Sub

Private Sub Ok_Button_Click()
    SaveSetting "ltxCrlEdt", "runtime", "header", Me.TextBox1.Value
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.TextBox1.Value = GetSetting("ltxCrlEdt", "runtime", "header", defhead)
End Sub

Private Sub UserForm_Terminate()
    Unload Me
End Sub
