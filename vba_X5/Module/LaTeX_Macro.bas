Attribute VB_Name = "LaTeX_Macro"
'VBA Macro for Corel Draw to easily import LaTeX-Output as vector Objects
'Macro is able to edit previously created LaTeX-Object while keeping position, scale, rotating and shearing
'Created by Jan Bender available at http://www.impulse-based.de/
'Modified for the use of Ghostscript and easy header-editing by Simon Klein (mail@simonklein.de)
'needs Ghostscript to be installed (path in registry or PATH-variable) and pdflatex in PATH (should already be with miktex and gpl ghostscript on windows)
'header stored in Registry

Sub LatexEdit()
    Dim frmEdit As New frmLatexEdit
    Dim s1 As Shape
    Set s1 = ActiveShape
    If s1 Is Nothing Then
        frmEdit.Show vbModal
    Else
        Load frmEdit
        On Error Resume Next
        Dim erg As String
        erg = ""
        erg = Dir(ActiveDocument.FilePath & s1.ObjectData("Comments"), vbNormal)
        If erg <> "" Then
            ' the comment is a filename
            Dim h As Integer
            h = FreeFile
            Open ActiveDocument.FilePath & s1.ObjectData("Comments") For Input As #h
            frmEdit.TextBox1.Text = Input$(LOF(h), h)
            Close #h
        Else
            ' the comment is directly the latex source code
            frmEdit.TextBox1.Text = s1.ObjectData("Comments")
        End If
        
        frmEdit.Show vbModal
    End If
End Sub

Function getGSPath() As String

  Const HKEY_LOCAL_MACHINE = &H80000002

  Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  Dim strkeypath As String
  strkeypath = "SOFTWARE"

  oReg.EnumKey HKEY_LOCAL_MACHINE, strkeypath, arrsubkeys
  If Not IsArray(arrsubkeys) Then Exit Function

  For Each strvalue In arrsubkeys
    Debug.Print strvalue
    If InStr(1, LCase(strvalue), "ghostscript") <> 0 Then

      oReg.EnumKey HKEY_LOCAL_MACHINE, strkeypath & "\" & strvalue, arrsubkeys1
      If IsArray(arrsubkeys1) Then
        For Each strvalue1 In arrsubkeys1
          'msg = strValue1 & vbCr

          oReg.GetStringValue HKEY_LOCAL_MACHINE, strkeypath & "\" & strvalue & "\" & strvalue1, "GS_DLL", GS_DLL
          'oReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath & "\" & strValue & "\" & strValue1, "GS_LIB", GS_LIB

          msg = GS_DLL
        Next
      End If
      'MsgBox msg
      getGSPath = Left(msg, InStrRev(msg, "\"))
    End If
  Next
End Function

Function getMiktexPath() As String
    Dim path As String
    path = Environ$("PATH")
    
    Dim DirS() As String
    Dim d As Variant
    DirS = Split(path, ";")
    
    For Each d In DirS
        If Dir(d + "\mgs.exe", vbNormal) <> "" Then
            'MsgBox D
            getMiktexPath = d
            Exit Function
        End If
    Next
End Function
