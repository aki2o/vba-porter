Attribute VB_Name = "main"
Option Explicit

Private lastmodified_of As Object
Private finished_preclose_of As Object
Private fso As Object

Public Sub initialize()
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub

Public Sub updateAll()
    If removeComponent = False Then
        MsgBox ""
        Exit Sub
    End If
End Sub

Public Function exportComponent() As Boolean
    Dim com As Object
    Dim expath As String
    Dim export As Boolean
    Dim confirmmsg As String
    Dim ret As Boolean

    exportComponent = True
    With ThisWorkbook.VBProject
        For Each com In 
End Function

Private Function removeComponent() As Boolean
End Function

