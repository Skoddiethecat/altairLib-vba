Attribute VB_Name = "alRunExporter"
'@Folder("AltairLib")
Option Explicit

Private Sub run()
    Dim AltairLib As AltairLib
    
    Set AltairLib = alFactory.AltairLibLoad
    
    AltairLib.alExporter.ExportVisualBasicCode

End Sub
