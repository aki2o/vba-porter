Attribute VB_Name = "main"
Option Explicit

Private lastmodified_of As Object
Private finished_preclose_of As Object
Private fso As Object

Public Sub initialize()
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set lastmodified_of = CreateObject("Scripting.Dictionary")
    Set finished_preclose_of = CreateObject("Scripting.Dictionary")
End Sub

Public Sub updateAll()
    If removeComponent = False Then
        MsgBox ""
        Exit Sub
    End If
End Sub

Public Function exportComponent() As Boolean
    Dim i As Integer
    Dim comnm As String
    Dim comtype As Variant
    Dim expath As String
    Dim export As Boolean
    Dim confirmmsg As String

    With ThisWorkbook.VBProject
        For i = 1 To .VBComponents.Count
            comnm = .VBComponents.Item(i).Name
            comtype = .VBComponents.Item(i).Type
            If isExportComponent(comnm, comtype) Then
                expath = getComponentPath(comnm)
                If expath <> "" Then

                    export = True
                    If isModified(expath) Then
                        export = False
                        confirmmsg = "以下のコンポーネントをエクスポートしようとしていますが、" & vbCrLf _
                                   & "エクスポート先のファイルが他ユーザによって変更されています。" & vbCrLf _
                                   & "このままエクスポートしてもよろしいですか？" & vbCrLf _
                                   & vbCrLf _
                                   & "コンポーネント名：" & comnm & vbCrLf _
                                   & "エクスポート先：" & expath
                        If MsgBox(confirmmsg, vbYesNo) = vbYes Then export = True
                    End If

                    If export Then
                        .VBComponents.Item(i).Export expath
                        updateModified expath
                    End If
                    
                End If
            End If
        Next
    End With
    
    exportComponent = True
End Function


'Subだと削除が完了する前に、次の処理が実行されてしまっているようなのでFunctionにした
Private Function removeComponent() As Boolean
    Dim i As Integer
    Dim comnm As String
    Dim comtype As Variant

    With ThisWorkbook.VBProject
        For i = 1 To .VBComponents.Count
            comnm = .VBComponents.Item(i).Name
            comtype = .VBComponents.Item(i).Type
            If isExportComponent(comnm, comtype) Then
                .VBComponents.Remove .VBComponents.Item(i)
            End If
        Next
    End With

    removeComponent = True
End Function

Private Sub importComponent(ByVal dirpath As String)
    Dim i As Integer
    Dim filenm As String
    Dim filepath As String

    If Not fso.FolderExists(dirpath) Then
        Exit Sub
    End If

    For i = 1 To fso.GetFolder(dirpath).Files.Count
        filenm = fso.GetFolder(dirpath).Files.Item(i).Name
        filepath = fso.GetFolder(dirpath).Files.Item(i).Path
        If isImportableFile(filenm) Then
            ThisWorkbook.VBProject.VBComponents.Import filepath
            updateModified filepath
            setupExportMetaInfo getComponentName(filenm), filepath
        End If
    Next
End Sub

Private Function isExportComponent(ByVal comnm As String, ByVal comtype As Variant) As Boolean
    If (comtype = 1 Or comtype = 2 Or comtype = 3) And comnm <> "main" Then isExportComponent = True
End Function

Private Function isImportableFile(ByVal filenm As String) As Boolean
End Function

Private Function getComponentName(ByVal filenm As String) As String
    Dim comnm As String
    
    comnm = filenm
    comnm = Replace(comnm, ".bas", "")
    comnm = Replace(comnm, ".cls", "")
    comnm = Replace(comnm, ".frm", "")
    getComponentName = comnm
End Function

Private Sub setupExportMetaInfo(ByVal comnm As String, ByVal exportpath As String)
    
End Sub

Private Function isModified(ByVal filepath As String) As Boolean
    Dim lastmodified As Variant

    lastmodified = lastmodified_of.Item(filepath)
    If Not lastmodified Then
        isModified = False
    ElseIf Not fso.FileExists(filepath) Then
        isModified = True
    ElseIf fso.GetFile(filepath).DateLastModified = lastmodified Then
        isModified = False
    Else
        isModified = True
    End If
End Function

Private Sub updateModified(ByVal filepath As String)
    If lastmodified_of.Exists(filepath) Then lastmodified_of.Remove filepath
    lastmodified_of.Add filepath, fso.GetFile(filepath).DateLastModified
End Sub

Private Function getComponentPath(ByVal comnm As String) As String
End Function

