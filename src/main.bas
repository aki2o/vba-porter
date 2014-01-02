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
    Dim com As Object
    Dim expath As String
    Dim export As Boolean
    Dim confirmmsg As String

    With ThisWorkbook.VBProject
        For Each com In .VBComponents
            If isExportComponent(com.Name, com.Type) Then
                expath = getExportPath(com.Name)
                If expath <> "" Then

                    export = True
                    If isModified(expath) Then
                        export = False
                        confirmmsg = "以下のコンポーネントをエクスポートしようとしていますが、" & vbCrLf _
                                   & "エクスポート先のファイルが他ユーザによって変更されています。" & vbCrLf _
                                   & "このままエクスポートしてもよろしいですか？" & vbCrLf _
                                   & vbCrLf _
                                   & "コンポーネント名：" & com.Name & vbCrLf _
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
    Dim com As Object

    With ThisWorkbook.VBProject
        For Each com In .VBComponents
            If isExportComponent(com.Name, com.Type) Then
                .VBComponents.Remove com
            End If
        Next
    End With

    removeComponent = True
End Function

Private Sub importComponent(ByVal dirpath As String)
    Dim f As Object
    Dim d As Object

    If Not fso.FolderExists(dirpath) Then Exit Sub

    For Each f In fso.GetFolder(dirpath).Files
        If isImportableFile(f.Name) Then
            ThisWorkbook.VBProject.VBComponents.Import f.Path
            updateModified filepath
            setExportPath getComponentName(f.Name), f.Path
        End If
    Next

    For Each d In fso.GetFolder(dirpath).SubFolders
        If isImportableFolder(d.Name) Then
            importComponent d.Path
        End If
    Next
End Sub

Private Function isExportComponent(ByVal comnm As String, ByVal comtype As Variant) As Boolean
    If (comtype = 1 Or comtype = 2 Or comtype = 3) And comnm <> "main" Then isExportComponent = True
End Function

Private Function isImportableFile(ByVal filenm As String) As Boolean
    If InStr(filenm, ".bas") > 0 Or _
       InStr(filenm, ".cls") > 0 Or _
       InStr(filenm, ".frm") > 0 Then
        isImportableFile = True
    End If
End Function

Private Function isImportableFolder(ByVal dirnm As String) As Boolean
    isImportableFolder = True
    If dirnm = ".svn" Then
        isImportableFolder = False
    End If
End Function

Private Function getComponentName(ByVal filenm As String) As String
    Dim comnm As String

    comnm = filenm
    comnm = Replace(comnm, ".bas", "")
    comnm = Replace(comnm, ".cls", "")
    comnm = Replace(comnm, ".frm", "")
    getComponentName = comnm
End Function


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


Private Sub setExportPath(ByVal comnm As String, ByVal exportpath As String)
    setMetaInfo comnm, "ExportPath", exportpath
End Sub

Private Function getExportPath(ByVal comnm As String) As String
    getExportPath = getMetaInfo(comnm, "ExportPath")
End Function


''''''''''''
' MetaInfo

Private Sub setMetaInfo(ByVal comnm As String, ByVal metanm As String, ByVal metavalue As String)
    Dim code As String
    Dim row As Long
    Dim currpath As String
    
    code = "'VBAPorter:" & metanm & "=" & metavalue
    With ThisWorkbook.VBProject.VBComponents.Item(comnm)
        currpath = getExportPath(comnm)
        If currpath = "" Then
            .CodeModule.InsertLines 1, code
        ElseIf currpath <> exportpath Then
            row = getMetaInfoRow(comnm, metanm)
            If row > 0 Then .CodeModule.ReplaceLine row, code
        End If
    End With
End Sub

Private Function getMetaInfo(ByVal comnm As String, ByVal metanm As String) As String
    Dim re As Object
    Dim row As Long
    Dim code As String

    row = getMetaInfoRow(comnm, metanm)
    If row <= 0 Then Exit Function
    
    Set re = newMetaInfoRegexp(metanm)
    With ThisWorkbook.VBProject.VBComponents.Item(comnm)
        code = .CodeModule.Lines(row, 1)
        If re.Test(code) Then getMetaInfo = re.Execute(code).Item(0).submatches(0)
    End With
End Function

Private Function getMetaInfoRow(ByVal comnm As String, ByVal metanm As String) As Long
    Dim re As Object
    Dim row As Long

    Set re = newMetaInfoRegexp(metanm)
    With ThisWorkbook.VBProject.VBComponents.Item(comnm)
        For row = 1 To .CodeModule.CountOfDeclarationLines
            If re.Test(.CodeModule.Lines(row, 1)) Then
                getMetaInfoRow = row
                Exit For
            End If
        Next
    End With
End Function

Private Function newMetaInfoRegexp(ByVal metanm As String) As Object
    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^'VBAPorter:" & metanm & "=(.+)$"
    re.IgnoreCase = True
    re.Global = True
    Set newMetaInfoRegexp = re
End Function

