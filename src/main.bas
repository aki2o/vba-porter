Attribute VB_Name = "main"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                         (ByVal lpApplicationName As String, _
                          ByVal lpKeyName As Any, _
                          ByVal lpDefault As String, _
                          ByVal lpReturnedString As String, _
                          ByVal nSize As Long, _
                          ByVal lpFileName As String) As Long

Private Const ROOTMENUNM As String = "VBAPorter"

Private fso As Object
Private lastmodified_of As Object
Private finished_preclose_of As Object

Public Sub initialize()
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set lastmodified_of = CreateObject("Scripting.Dictionary")
    Set finished_preclose_of = CreateObject("Scripting.Dictionary")
End Sub

Public Sub updateAll()
    Dim rootmenu As CommandBarPopup
    Dim configs() As String
    Dim i As Integer
    Dim dirpath As String

    On Error GoTo CATCH_ERR
        
    deleteMenu
    Set rootmenu = createMenu
    
    If removeComponent = False Then
        MsgBox "既存コンポーネントの削除に失敗しました"
        Exit Sub
    End If
    
    If Not existConfigFile Then
        MsgBox getConfigFolderPath & " に設定ファイルが見つかりません"
        Exit Sub
    End If
    configs = getConfigList()
    For i = 0 To UBound(configs)
    
        On Error GoTo FAILED_SETUP
        
        dirpath = getConfigValue(configs(i), "ROOT")
        If Not fso.FolderExists(dirpath) Then
            MsgBox "以下のフォルダが見つからないため、" & vbCrLf _
                   & configs(i) & "のコンポーネントのインポート及びメニュー生成は実行されません。" & vbCrLf _
                   & vbCrLf _
                   & "フォルダ：" & dirpath
        Else
            importComponent dirpath
            createMenuFromFolder rootmenu, dirpath
        End If
        GoTo NEXT_LOOP

FAILED_SETUP:
        MsgBox configs(i) & "のコンポーネントのインポート及びメニュー生成に失敗しました。" & vbCrLf _
               & vbCrLf _
               & "ErrNumber: " & Err.Number & vbCrLf _
               & "ErrSource: " & Err.Source & vbCrLf _
               & Err.Description
NEXT_LOOP:

    Next

    MsgBox "完了しました"
    Exit Sub

CATCH_ERR:
    MsgBox "実行に失敗しました。" & vbCrLf _
           & vbCrLf _
           & "ErrNumber: " & Err.Number & vbCrLf _
           & "ErrSource: " & Err.Source & vbCrLf _
           & Err.Description
End Sub

Public Sub save()

    On Error GoTo CATCH_ERR
    
    exportComponent
    MsgBox "完了しました"
    Exit Sub
    
CATCH_ERR:
    MsgBox "実行に失敗しました。" & vbCrLf _
           & vbCrLf _
           & "ErrNumber: " & Err.Number & vbCrLf _
           & "ErrSource: " & Err.Source & vbCrLf _
           & Err.Description
End Sub


''''''''
' Menu

Public Function createMenu() As CommandBarPopup
    Dim bar As CommandBar
    Dim rootmenu As CommandBarPopup
    Dim childmenu As CommandBarPopup
    Dim menubtn As CommandBarButton

    On Error GoTo CATCH_ERR
    
    Set bar = Application.CommandBars("Worksheet Menu Bar")
    Set rootmenu = bar.Controls.Add(Type:=MsoControlType.msoControlPopup)
    rootmenu.Caption = ROOTMENUNM
    
    Set childmenu = rootmenu.Controls.Add(Type:=MsoControlType.msoControlPopup)
    childmenu.Caption = "管理"
    Set menubtn = childmenu.Controls.Add(Type:=MsoControlType.msoControlButton)
    menubtn.Caption = "保存"
    menubtn.OnAction = "main.save"
    Set menubtn = childmenu.Controls.Add(Type:=MsoControlType.msoControlButton)
    menubtn.Caption = "全て更新"
    menubtn.OnAction = "main.updateAll"

    Set createMenu = rootmenu
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "createMenu > " & Err.Source, Err.Description
End Function

Public Sub deleteMenu()
    Dim bar As CommandBar
    Dim i As Integer

    On Error GoTo CATCH_ERR
    
    Set bar = Application.CommandBars("Worksheet Menu Bar")
    For i = 1 To bar.Controls.Count
        If bar.Controls.Item(i).Caption = ROOTMENUNM Then
            bar.Controls.Item(i).Delete
            Exit For
        End If
    Next
    Exit Sub

CATCH_ERR:
    Err.Raise Err.Number, "deleteMenu > " & Err.Source, Err.Description
End Sub

Private Sub createMenuFromFolder(ByRef parent As CommandBarPopup, ByVal dirpath As String)
    Dim d As Object
    Dim f As Object
    Dim menu As CommandBarPopup
    Dim comnm As String
    Dim btnnm As String
    Dim btn As CommandBarButton

    On Error GoTo CATCH_ERR
    
    Set menu = parent.Controls.Add(Type:=MsoControlType.msoControlPopup)
    menu.Caption = fso.GetFolder(dirpath).Name
    
    For Each d In fso.GetFolder(dirpath).SubFolders
        If isImportableFolder(d.Name) Then
            On Error GoTo FAILED_RECURSIVE
            createMenuFromFolder menu, d.Path
            On Error GoTo CATCH_ERR
        End If
    Next
    
    For Each f In fso.GetFolder(dirpath).Files
        If isMenuableFile(f.Name) Then
            comnm = getComponentName(f.Name)
            btnnm = getMetaInfo(comnm, "MenuName")
            If btnnm <> "" Then
                Set btn = menu.Controls.Add(Type:=MsoControlType.msoControlButton)
                btn.Caption = btnnm
                btn.OnAction = comnm & ".Click"
            End If
        End If
    Next
    Exit Sub

FAILED_RECURSIVE:
    Err.Raise Err.Number, Err.Source, Err.Description
CATCH_ERR:
    Err.Raise Err.Number, "createMenuFromFolder > " & Err.Source, Err.Description
End Sub

Private Function isMenuableFile(ByVal filenm As String) As Boolean
    On Error GoTo CATCH_ERR
    
    If InStr(filenm, ".bas") > 0 Then
        isMenuableFile = True
    End If
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "isMenuableFile > " & Err.Source, Err.Description
End Function


'''''''''''''
' Component

Private Sub importComponent(ByVal dirpath As String)
    Dim f As Object
    Dim d As Object

    On Error GoTo CATCH_ERR
    
    For Each f In fso.GetFolder(dirpath).Files
        If isImportableFile(f.Name) Then
            ThisWorkbook.VBProject.VBComponents.Import f.Path
            updateModified f.Path
            setMetaInfo getComponentName(f.Name), "ExportPath", f.Path
        End If
    Next
    For Each d In fso.GetFolder(dirpath).SubFolders
        If isImportableFolder(d.Name) Then
            On Error GoTo FAILED_RECURSIVE
            importComponent d.Path
            On Error GoTo CATCH_ERR
        End If
    Next
    Exit Sub

FAILED_RECURSIVE:
    Err.Raise Err.Number, Err.Source, Err.Description
CATCH_ERR:
    Err.Raise Err.Number, "importComponent > " & Err.Source, Err.Description
End Sub

Private Sub exportComponent()
    Dim com As Object
    Dim expath As String
    Dim export As Boolean
    Dim confirmmsg As String

    On Error GoTo CATCH_ERR
    
    With ThisWorkbook.VBProject
        For Each com In .VBComponents
            If isExportComponent(com.Name, com.Type) Then
                expath = getMetaInfo(com.Name, "ExportPath")
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
                        com.export expath
                        updateModified expath
                    End If
                    
                End If
            End If
        Next
    End With
    Exit Sub

CATCH_ERR:
    Err.Raise Err.Number, "exportComponent > " & Err.Source, Err.Description
End Sub

'Subだと削除が完了する前に、次の処理が実行されてしまっているようなのでFunctionにした
Private Function removeComponent() As Boolean
    Dim com As Object

    On Error GoTo CATCH_ERR
    
    With ThisWorkbook.VBProject
        For Each com In .VBComponents
            If isExportComponent(com.Name, com.Type) Then
                .VBComponents.Remove com
            End If
        Next
    End With

    removeComponent = True
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "removeComponent > " & Err.Source, Err.Description
End Function

Private Function isImportableFile(ByVal filenm As String) As Boolean
    On Error GoTo CATCH_ERR
    
    If InStr(filenm, ".bas") > 0 Or _
       InStr(filenm, ".cls") > 0 Or _
       InStr(filenm, ".frm") > 0 Then
        isImportableFile = True
    End If
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "isImportableFile > " & Err.Source, Err.Description
End Function

Private Function isImportableFolder(ByVal dirnm As String) As Boolean
    On Error GoTo CATCH_ERR
    
    isImportableFolder = True
    If dirnm = ".svn" Then
        isImportableFolder = False
    End If
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "isImportableFolder > " & Err.Source, Err.Description
End Function

Private Function isExportComponent(ByVal comnm As String, ByVal comtype As Variant) As Boolean
    On Error GoTo CATCH_ERR
    
    If (comtype = 1 Or comtype = 2 Or comtype = 3) And comnm <> "main" Then isExportComponent = True
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "isExportComponent > " & Err.Source, Err.Description
End Function

Private Function getComponentName(ByVal filenm As String) As String
    Dim comnm As String

    On Error GoTo CATCH_ERR
    
    comnm = filenm
    comnm = Replace(comnm, ".bas", "")
    comnm = Replace(comnm, ".cls", "")
    comnm = Replace(comnm, ".frm", "")
    getComponentName = comnm
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "getComponentName > " & Err.Source, Err.Description
End Function


'''''''''''''''''''
' Manage Modified

Private Function isModified(ByVal filepath As String) As Boolean
    Dim lastmodified As Variant

    On Error GoTo CATCH_ERR
    
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
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "isModified > " & Err.Source, Err.Description
End Function

Private Sub updateModified(ByVal filepath As String)
    On Error GoTo CATCH_ERR
    
    If lastmodified_of.Exists(filepath) Then lastmodified_of.Remove filepath
    lastmodified_of.Add filepath, fso.GetFile(filepath).DateLastModified
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "updateModified > " & Err.Source, Err.Description
End Sub


''''''''''''
' MetaInfo

Private Sub setMetaInfo(ByVal comnm As String, ByVal metanm As String, ByVal metavalue As String)
    Dim code As String
    Dim row As Long
    
    On Error GoTo CATCH_ERR
    
    code = buildMetaInfoCode(metanm, metavalue)
    With ThisWorkbook.VBProject.VBComponents.Item(comnm)
        row = getMetaInfoRow(comnm, metanm)
        If row > 0 Then
            .CodeModule.ReplaceLine row, code
        Else
            .CodeModule.InsertLines 1, code
        End If
    End With
    Exit Sub

CATCH_ERR:
    Err.Raise Err.Number, "setMetaInfo > " & Err.Source, Err.Description
End Sub

Private Function getMetaInfo(ByVal comnm As String, ByVal metanm As String) As String
    Dim re As Object
    Dim row As Long
    Dim code As String

    On Error GoTo CATCH_ERR
    
    row = getMetaInfoRow(comnm, metanm)
    If row <= 0 Then Exit Function
    
    Set re = newMetaInfoRegexp(metanm)
    With ThisWorkbook.VBProject.VBComponents.Item(comnm)
        code = .CodeModule.Lines(row, 1)
        If re.test(code) Then getMetaInfo = re.Execute(code).Item(0).submatches(0)
    End With
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "getMetaInfo > " & Err.Source, Err.Description
End Function

Private Function getMetaInfoRow(ByVal comnm As String, ByVal metanm As String) As Long
    Dim re As Object
    Dim row As Long

    On Error GoTo CATCH_ERR
    
    Set re = newMetaInfoRegexp(metanm)
    With ThisWorkbook.VBProject.VBComponents.Item(comnm)
        For row = 1 To .CodeModule.CountOfDeclarationLines
            If re.test(.CodeModule.Lines(row, 1)) Then
                getMetaInfoRow = row
                Exit For
            End If
        Next
    End With
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "getMetaInfoRow > " & Err.Source, Err.Description
End Function

Private Function buildMetaInfoCode(ByVal metanm As String, ByVal metavalue As String) As String
    On Error GoTo CATCH_ERR
    
    buildMetaInfoCode = "'VBAPorter:" & metanm & "=" & metavalue
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "buildMetaInfoCode > " & Err.Source, Err.Description
End Function

Private Function newMetaInfoRegexp(ByVal metanm As String) As Object
    Dim re As Object

    On Error GoTo CATCH_ERR
    
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^'VBAPorter:" & metanm & "=(.+)$"
    re.IgnoreCase = True
    re.Global = True
    Set newMetaInfoRegexp = re
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "newMetaInfoRegexp > " & Err.Source, Err.Description
End Function


''''''''''
' Config

Private Function existConfigFile() As Boolean
    On Error GoTo CATCH_ERR
    
    existConfigFile = fso.FileExists(getConfigFilePath)
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "existConfigFile > " & Err.Source, Err.Description
End Function

Private Function getConfigList() As String()
    Dim buff As String * 32767
    Dim retcd As Long
    Dim retvalue As String

    On Error GoTo CATCH_ERR
    
    retcd = GetPrivateProfileString(vbNullString, vbNullString, vbNullString, buff, Len(buff), getConfigFilePath)
    If retcd = 0 Then
        Exit Function
    End If
    retvalue = Strings.Left(buff, InStr(buff, vbNullChar & vbNullChar) - 1)
    getConfigList = Split(retvalue, vbNullChar)
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "getConfigList > " & Err.Source, Err.Description
End Function

Private Function getConfigValue(ByVal section As String, ByVal key As String) As String
    Dim buff As String * 32767
    Dim retcd As Long

    On Error GoTo CATCH_ERR
    
    retcd = GetPrivateProfileString(section, key, "", buff, Len(buff), getConfigFilePath)
    If retcd = 0 Then
        Exit Function
    End If
    getConfigValue = Strings.Left(buff, InStr(buff, vbNullChar) - 1)
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "getConfigValue > " & Err.Source, Err.Description
End Function

Private Function getConfigFilePath() As String
    Dim dirpath As String
    
    On Error GoTo CATCH_ERR
    
    dirpath = getConfigFolderPath
    If fso.FileExists(dirpath & "\_vbaporter") Then
        getConfigFilePath = dirpath & "\_vbaporter"
    Else
        getConfigFilePath = dirpath & "\.vbaporter"
    End If
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "getConfigFilePath > " & Err.Source, Err.Description
End Function

Private Function getConfigFolderPath() As String
    Dim homedir As String

    On Error GoTo CATCH_ERR
    
    homedir = Environ("HOME")
    If Not fso.FolderExists(homedir) Then homedir = Environ("USERPROFILE")
    getConfigFolderPath = homedir
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "getConfigFolderPath > " & Err.Source, Err.Description
End Function

