Attribute VB_Name = "modInstaller"
' ════════════════════════════════════════════════════════════════════════
'  Curriculum Map Builder — code installer
'  Loads (or reloads) every other module from the "vba" folder next to this
'  workbook, so the code can be updated without copying it in by hand.
'
'  One-time setup (see ../SETUP.md):
'    Excel ▸ Preferences/Options ▸ Security/Trust Center ▸ tick
'    "Trust access to the VBA project object model".
'  Then run:  ReloadModules   (Tools ▸ Macro ▸ Macros…, or Alt/Opt+F8)
' ════════════════════════════════════════════════════════════════════════
Option Explicit

Private Const INSTALLER As String = "modInstaller"

Public Sub ReloadModules()
    Dim sep As String, folder As String, f As String, n As Long
    Dim comp As Object, toRemove As Collection, item As Variant

    sep = Application.PathSeparator
    folder = ThisWorkbook.Path & sep & "vba" & sep
    If Len(Dir(folder & INSTALLER & ".bas")) = 0 Then
        MsgBox "Can't find the code folder:" & vbLf & folder & vbLf & vbLf & _
               "Keep this workbook in the Curriculum-Map-Builder folder, next to its 'vba' folder.", vbExclamation
        Exit Sub
    End If

    On Error GoTo NoAccess
    ' 1. remove standard modules and classes (not this installer)
    Set toRemove = New Collection
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If (comp.Type = 1 Or comp.Type = 2) And comp.Name <> INSTALLER Then toRemove.Add comp
    Next comp
    For Each item In toRemove
        ThisWorkbook.VBProject.VBComponents.Remove item
    Next item

    ' 2. import every .bas / .cls in the folder (Dir without wildcards: Mac-safe)
    f = Dir(folder)
    Do While Len(f) > 0
        If (LCase$(Right$(f, 4)) = ".bas" Or LCase$(Right$(f, 4)) = ".cls") _
           And LCase$(f) <> LCase$(INSTALLER & ".bas") Then
            ThisWorkbook.VBProject.VBComponents.Import folder & f
            n = n + 1
        End If
        f = Dir()
    Loop

    ' 3. workbook event code (GLO clicks etc.) goes into ThisWorkbook
    If Len(Dir(folder & "ThisWorkbook.txt")) > 0 Then
        With ThisWorkbook.VBProject.VBComponents(ThisWorkbook.CodeName).CodeModule
            If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
            .AddFromFile folder & "ThisWorkbook.txt"
        End With
    End If

    Application.StatusBar = False
    MsgBox "Loaded " & n & " module" & IIf(n = 1, "", "s") & " from" & vbLf & folder, vbInformation
    Exit Sub

NoAccess:
    MsgBox "Excel blocked access to the code (" & Err.Description & ")." & vbLf & vbLf & _
           "Turn on: Preferences/Options ▸ Security/Trust Center ▸ " & _
           """Trust access to the VBA project object model"", then run ReloadModules again.", vbExclamation
End Sub

' Quiet variant for automated reloads (no message boxes)
Public Function ReloadModulesQuiet() As String
    Application.DisplayAlerts = False
    On Error GoTo Fail
    Dim sep As String, folder As String, f As String, n As Long
    Dim comp As Object, toRemove As Collection, item As Variant
    sep = Application.PathSeparator
    folder = ThisWorkbook.Path & sep & "vba" & sep
    Set toRemove = New Collection
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If (comp.Type = 1 Or comp.Type = 2) And comp.Name <> INSTALLER Then toRemove.Add comp
    Next comp
    For Each item In toRemove
        ThisWorkbook.VBProject.VBComponents.Remove item
    Next item
    f = Dir(folder)
    Do While Len(f) > 0
        If (LCase$(Right$(f, 4)) = ".bas" Or LCase$(Right$(f, 4)) = ".cls") _
           And LCase$(f) <> LCase$(INSTALLER & ".bas") Then
            ThisWorkbook.VBProject.VBComponents.Import folder & f
            n = n + 1
        End If
        f = Dir()
    Loop
    If Len(Dir(folder & "ThisWorkbook.txt")) > 0 Then
        With ThisWorkbook.VBProject.VBComponents(ThisWorkbook.CodeName).CodeModule
            If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
            .AddFromFile folder & "ThisWorkbook.txt"
        End With
    End If
    ReloadModulesQuiet = "OK " & n
    Exit Function
Fail:
    ReloadModulesQuiet = "ERR " & Err.Number & " " & Err.Description
End Function
