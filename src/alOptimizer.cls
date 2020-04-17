VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "alOptimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("AltairLib.Modules")
Option Explicit

Public Active As Boolean

Public Sub Enable()
    With Application
        .ScreenUpdating = False
        .DisplayStatusBar = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .PrintCommunication = False
    
    End With
    
    ActiveSheet.DisplayPageBreaks = False 'this is a sheet-level setting
    
    Active = True

End Sub

Public Sub Disable()
    With Application
        .ScreenUpdating = True
        .DisplayStatusBar = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .PrintCommunication = True
    
    End With
    
    ActiveSheet.DisplayPageBreaks = True
    
    Active = False

End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Active Then Disable

End Sub

