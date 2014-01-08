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

Private Enum lang
    English = 1
    Japanese = 2
End Enum

Private Enum msg
    FINISHED = 1
    FAILED = 11
    FAILED_MENU = 12
    FAILED_IMPORT = 13
    FAILED_REMOVE = 14
    MENU_MANAGE = 21
    MENU_EXPORT = 22
    MENU_IMPORT = 23
    CONFIRM_EXPORT = 31
    NONE_EXPORTPATH = 41
    NONE_CONFIG = 42
    NONE_ROOTPATH = 43
    INFO_FOLDER = 51
    INFO_FILE = 52
End Enum

Public Sub initialize(Optional ByVal quiet As Boolean = False)
    On Error GoTo CATCH_ERR
        
    deleteMenu
    createMenu
    If Not quiet Then popupFinish
    Exit Sub
    
CATCH_ERR:
    popupError Err.Number, Err.Source, Err.Description
End Sub

Public Sub finalize(Optional ByVal quiet As Boolean = False)
    On Error GoTo CATCH_ERR
        
    Set fso = CreateObject("Scripting.FileSystemObject")
    deleteMenu
    removeModified
    If Not quiet Then popupFinish
    Exit Sub
    
CATCH_ERR:
    popupError Err.Number, Err.Source, Err.Description
End Sub

Public Sub update()
    Dim rootmenu As CommandBarPopup
    Dim configs() As String
    Dim i As Integer
    Dim dirpath As String
    Dim rootdirs() As String
    Dim menunms() As String

    On Error GoTo CATCH_ERR

    ' Init
    Set fso = CreateObject("Scripting.FileSystemObject")
    deleteMenu
    Set rootmenu = createMenu
    If removeComponent = False Then
        popupMsg getMsg(FAILED_REMOVE)
        Exit Sub
    End If

    ' Get config
    If Not existConfigFile Then
        popupMsg formatString(getMsg(NONE_CONFIG), getConfigFolderPath)
        Exit Sub
    End If
    configs = getConfigList()

    ' Check config
    ReDim rootdirs(0)
    ReDim menunms(0)
    For i = 0 To UBound(configs)
        dirpath = getConfigValue(configs(i), "ROOT")
        If Not fso.FolderExists(dirpath) Then
            popupMsg formatString(getMsg(NONE_ROOTPATH), configs(i), dirpath)
        Else
            ReDim Preserve rootdirs(UBound(rootdirs) + 1)
            rootdirs(UBound(rootdirs)) = dirpath
            ReDim Preserve menunms(UBound(menunms) + 1)
            menunms(UBound(menunms)) = getConfigValue(configs(i), "MENUNAME")
        End If
    Next

    ' Import
    For i = 1 To UBound(rootdirs)
        On Error GoTo FAILED_IMPORT
        importComponent rootdirs(i)
        GoTo NEXT_IMPORT
FAILED_IMPORT:
        popupError Err.Number, Err.Source, Err.Description, getMsg(FAILED_IMPORT)
NEXT_IMPORT:
    Next

    ' Create menu
    For i = 1 To UBound(menunms)
        On Error GoTo FAILED_MENU
        If menunms(i) <> "" Then createMenuFromFolder rootmenu, rootdirs(i), menunms(i)
        GoTo NEXT_MENU
FAILED_MENU:
        popupError Err.Number, Err.Source, Err.Description, getMsg(FAILED_MENU)
NEXT_MENU:
    Next

    On Error GoTo CATCH_ERR
    
    ' Save information of modified
    For i = 1 To UBound(rootdirs)
        updateModifiedRecursive rootdirs(i)
    Next
    saveModified

    popupFinish
    Exit Sub

CATCH_ERR:
    popupError Err.Number, Err.Source, Err.Description
End Sub

Public Sub save()
    On Error GoTo CATCH_ERR
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    loadModified
    exportComponent
    saveModified
    popupFinish
    Exit Sub
    
CATCH_ERR:
    popupError Err.Number, Err.Source, Err.Description
End Sub


''''''''
' Menu

Private Function createMenu() As CommandBarPopup
    Dim bar As CommandBar
    Dim rootmenu As CommandBarPopup
    Dim childmenu As CommandBarPopup
    Dim menubtn As CommandBarButton

    On Error GoTo CATCH_ERR
    
    Set bar = Application.CommandBars("Worksheet Menu Bar")
    Set rootmenu = bar.Controls.Add(Type:=MsoControlType.msoControlPopup)
    rootmenu.Caption = ROOTMENUNM
    
    Set childmenu = rootmenu.Controls.Add(Type:=MsoControlType.msoControlPopup)
    childmenu.Caption = getMsg(MENU_MANAGE)
    Set menubtn = childmenu.Controls.Add(Type:=MsoControlType.msoControlButton)
    menubtn.Caption = getMsg(MENU_EXPORT)
    menubtn.OnAction = "main.save"
    Set menubtn = childmenu.Controls.Add(Type:=MsoControlType.msoControlButton)
    menubtn.Caption = getMsg(MENU_IMPORT)
    menubtn.OnAction = "main.update"

    Set createMenu = rootmenu
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "createMenu > " & Err.Source, Err.Description
End Function

Private Sub deleteMenu()
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

Private Sub createMenuFromFolder(ByRef parent As CommandBarPopup, _
                                 ByVal dirpath As String, _
                                 Optional ByVal menunm As String)
    Dim d As Object
    Dim f As Object
    Dim menu As CommandBarPopup
    Dim com As Object
    Dim btnnm As String
    Dim btn As CommandBarButton
    Dim errmsg As String

    On Error GoTo CATCH_ERR
    
    Set menu = parent.Controls.Add(Type:=MsoControlType.msoControlPopup)
    If menunm = "" Then menunm = fso.GetFolder(dirpath).Name
    menu.Caption = menunm
    
    For Each d In fso.GetFolder(dirpath).SubFolders
        If isImportableFolder(d.Name) Then
            On Error GoTo FAILED_RECURSIVE
            createMenuFromFolder menu, d.Path
            On Error GoTo CATCH_ERR
        End If
    Next
    
    For Each f In fso.GetFolder(dirpath).Files
        If isMenuableFile(f.Name) Then
            Set com = ThisWorkbook.VBProject.VBComponents.Item(getComponentName(f.Name))
            If Not com Is Nothing Then
                btnnm = getMetaInfo(com.Name, "MenuName")
                If btnnm <> "" Then
                    Set btn = menu.Controls.Add(Type:=MsoControlType.msoControlButton)
                    btn.Caption = btnnm
                    btn.OnAction = com.Name & ".Click"
                End If
            End If
        End If
    Next
    Exit Sub

FAILED_RECURSIVE:
    Err.Raise Err.Number, Err.Source, Err.Description
CATCH_ERR:
    If Not d Is Nothing Then errmsg = formatString(getMsg(INFO_FOLDER), d.Path)
    If Not f Is Nothing Then errmsg = formatString(getMsg(INFO_FILE), f.Path)
    Err.Raise Err.Number, "createMenuFromFolder > " & Err.Source, errmsg & Err.Description
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
    Dim com As Object
    Dim d As Object
    Dim errmsg As String

    On Error GoTo CATCH_ERR
    
    For Each f In fso.GetFolder(dirpath).Files
        If isImportableFile(f.Name) Then
            Set com = ThisWorkbook.VBProject.VBComponents.Import(f.Path)
            setMetaInfo com.Name, "ExportPath", f.Path
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
    If Not d Is Nothing Then errmsg = formatString(getMsg(INFO_FOLDER), d.Path)
    If Not f Is Nothing Then errmsg = formatString(getMsg(INFO_FILE), f.Path)
    Err.Raise Err.Number, "importComponent > " & Err.Source, errmsg & Err.Description
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
                If expath = "" Then
                    popupMsg formatString(getMsg(NONE_EXPORTPATH), com.Name)
                Else
                    export = True
                    If isModified(expath) Then
                        export = False
                        confirmmsg = formatString(getMsg(CONFIRM_EXPORT), com.Name, expath)
                        If popupMsg(confirmmsg, vbYesNo) = vbYes Then export = True
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

'Sub���ƍ폜����������O�ɁA���̏��������s����Ă��܂��Ă���悤�Ȃ̂�Function�ɂ���
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

Private Sub loadModified()
    Dim filepath As String
    Dim mgr As Object
    Dim elem() As String

    On Error GoTo CATCH_ERR
    
    filepath = getModifiedStorePath
    If Not fso.FileExists(filepath) Then Exit Sub
    Set mgr = getModifiedManager
    With fso.OpenTextFile(filepath, 1, False, -2)
        Do While Not .AtEndOfStream
            elem = Split(.ReadLine, vbTab)
            If UBound(elem) = 1 Then
                If mgr.Exists(elem(0)) Then mgr.Remove elem(0)
                mgr.Add elem(0), elem(1)
            End If
        Loop
        .Close
    End With
    Exit Sub

CATCH_ERR:
    Err.Raise Err.Number, "loadModified > " & Err.Source, Err.Description
End Sub

Private Sub saveModified()
    Dim filepath As String
    Dim mgr As Object
    Dim key As Variant
    
    On Error GoTo CATCH_ERR
    
    filepath = getModifiedStorePath
    Set mgr = getModifiedManager
    With fso.OpenTextFile(filepath, 2, True, -2)
        For Each key In mgr.Keys
            .WriteLine key & vbTab & mgr.Item(key)
        Next
        .Close
    End With
    Exit Sub

CATCH_ERR:
    Err.Raise Err.Number, "saveModified > " & Err.Source, Err.Description
End Sub

Private Sub removeModified()
    Dim filepath As String
    
    On Error GoTo CATCH_ERR
    
    filepath = getModifiedStorePath
    If Not fso.FileExists(filepath) Then Exit Sub
    fso.DeleteFile filepath
    Exit Sub

CATCH_ERR:
    Err.Raise Err.Number, "removeModified > " & Err.Source, Err.Description
End Sub

Private Sub updateModifiedRecursive(ByVal dirpath As String)
    Dim f As Object
    Dim d As Object
    Dim errmsg As String

    On Error GoTo CATCH_ERR
    
    For Each f In fso.GetFolder(dirpath).Files
        If isImportableFile(f.Name) Then updateModified f.Path
    Next
    For Each d In fso.GetFolder(dirpath).SubFolders
        If isImportableFolder(d.Name) Then
            On Error GoTo FAILED_RECURSIVE
            updateModifiedRecursive d.Path
            On Error GoTo CATCH_ERR
        End If
    Next
    Exit Sub

FAILED_RECURSIVE:
    Err.Raise Err.Number, Err.Source, Err.Description
CATCH_ERR:
    If Not d Is Nothing Then errmsg = formatString(getMsg(INFO_FOLDER), d.Path)
    If Not f Is Nothing Then errmsg = formatString(getMsg(INFO_FILE), f.Path)
    Err.Raise Err.Number, "updateModifiedRecursive > " & Err.Source, errmsg & Err.Description
End Sub

Private Function isModified(ByVal filepath As String) As Boolean
    Dim mgr As Object
    Dim currvalue As String
    Dim storevalue As String

    On Error GoTo CATCH_ERR
    
    If Not fso.FileExists(filepath) Then Exit Function
    Set mgr = getModifiedManager
    If Not mgr.Exists(filepath) Then Exit Function
    currvalue = fso.GetFile(filepath).DateLastModified
    storevalue = mgr.Item(filepath)
    If currvalue = storevalue Then Exit Function
    isModified = True
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "isModified > " & Err.Source, Err.Description
End Function

Private Sub updateModified(ByVal filepath As String)
    Dim mgr As Object
    
    On Error GoTo CATCH_ERR
    
    Set mgr = getModifiedManager
    If mgr.Exists(filepath) Then mgr.Remove filepath
    mgr.Add filepath, fso.GetFile(filepath).DateLastModified
    Exit Sub

CATCH_ERR:
    Err.Raise Err.Number, "updateModified > " & Err.Source, Err.Description
End Sub

Private Function getModifiedManager() As Object
    Static ret As Object
    
    On Error GoTo CATCH_ERR
    
    If ret Is Nothing Then Set ret = CreateObject("Scripting.Dictionary")
    Set getModifiedManager = ret
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "getModifiedManager > " & Err.Source, Err.Description
End Function

Private Function getModifiedStorePath() As String
    On Error GoTo CATCH_ERR
    
    getModifiedStorePath = fso.GetSpecialFolder(2) & "\vbaporter.modified"
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "getModifiedStorePath > " & Err.Source, Err.Description
End Function


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


''''''''''''''''
' Notification

Private Sub popupFinish(Optional ByVal msg As String)
    If msg = "" Then msg = getMsg(FINISHED)
    popupMsg msg
End Sub

Private Sub popupError(ByVal errno As Long, _
                       ByVal errsrc As String, _
                       ByVal errdesc As String, _
                       Optional ByVal errmsg As String)
    If errmsg = "" Then errmsg = getMsg(FAILED)
    popupMsg errmsg & vbCrLf _
             & vbCrLf _
             & "ErrNumber: " & errno & vbCrLf _
             & "ErrSource: " & errsrc & vbCrLf _
             & errdesc
End Sub

Private Function popupMsg(ByVal msg As String, Optional ByVal style As VbMsgBoxStyle = vbOKOnly) As VbMsgBoxResult
    popupMsg = MsgBox(msg, style, "VBAPorter")
End Function

Private Function formatString(ByVal s As String, ParamArray args() As Variant) As String
    Dim i As Integer
    Dim idx As Long

    On Error GoTo CATCH_ERR
    
    For i = 0 To UBound(args)
        idx = InStr(s, "%s")
        If idx <= 0 Then Exit For
        s = Left(s, idx - 1) & args(i) & Mid(s, idx + 2)
    Next
    formatString = s
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "formatString > " & Err.Source, Err.Description
End Function

Private Function getMsg(ByVal msgtype As msg) As String
    Dim key As String
    Dim mgr As Object

    On Error GoTo CATCH_ERR
    
    key = getMsgKey(msgtype, getLanguage)
    Set mgr = getMsgManager
    If Not mgr.Exists(key) Then Exit Function
    getMsg = mgr.Item(key)
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "getMsg > " & Err.Source, Err.Description
End Function

Private Function getLanguage() As lang
    On Error GoTo CATCH_ERR
    
    Select Case Application.International(xlCountryCode)
        Case 81
            getLanguage = lang.Japanese
        Case Else
            getLanguage = lang.English
    End Select
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "getLanguage > " & Err.Source, Err.Description
End Function

Private Function getMsgManager() As Object
    Static ret As Object
    
    On Error GoTo CATCH_ERR
    
    If ret Is Nothing Then
        Set ret = CreateObject("Scripting.Dictionary")
        
        ret.Add getMsgKey(msg.FINISHED, lang.English), _
                "Finished execute."
        ret.Add getMsgKey(msg.FINISHED, lang.Japanese), _
                "���s���������܂����B"
        
        ret.Add getMsgKey(msg.FAILED, lang.English), _
                "Failed to execute."
        ret.Add getMsgKey(msg.FAILED, lang.Japanese), _
                "���s�Ɏ��s���܂����B"
        
        ret.Add getMsgKey(msg.FAILED_MENU, lang.English), _
                "Failed to create menu."
        ret.Add getMsgKey(msg.FAILED_MENU, lang.Japanese), _
                "���j���[�����Ɏ��s���܂����B"
        
        ret.Add getMsgKey(msg.FAILED_IMPORT, lang.English), _
                "Failed to import component."
        ret.Add getMsgKey(msg.FAILED_IMPORT, lang.Japanese), _
                "�R���|�[�l���g�̃C���|�[�g�Ɏ��s���܂����B"
        
        ret.Add getMsgKey(msg.FAILED_REMOVE, lang.English), _
                "Failed to remove component."
        ret.Add getMsgKey(msg.FAILED_REMOVE, lang.Japanese), _
                "�����R���|�[�l���g�̍폜�Ɏ��s���܂����B"
        
        ret.Add getMsgKey(msg.MENU_MANAGE, lang.English), _
                "Management"
        ret.Add getMsgKey(msg.MENU_MANAGE, lang.Japanese), _
                "�Ǘ�"

        ret.Add getMsgKey(msg.MENU_EXPORT, lang.English), _
                "Export"
        ret.Add getMsgKey(msg.MENU_EXPORT, lang.Japanese), _
                "�G�N�X�|�[�g"

        ret.Add getMsgKey(msg.MENU_IMPORT, lang.English), _
                "Import"
        ret.Add getMsgKey(msg.MENU_IMPORT, lang.Japanese), _
                "�C���|�[�g"
        
        ret.Add getMsgKey(msg.CONFIRM_EXPORT, lang.English), _
                "The file as a export path of the following component was updated by other user." & vbCrLf _
                & "Do you continue to export the component?" & vbCrLf _
                & vbCrLf _
                & "Component: %s" & vbCrLf _
                & "ExportPath: %s"
        ret.Add getMsgKey(msg.CONFIRM_EXPORT, lang.Japanese), _
                "�ȉ��̃R���|�[�l���g�̓G�N�X�|�[�g��̃t�@�C���������[�U�ɂ���ĕύX����Ă��܂��B" & vbCrLf _
                & "���̂܂܃G�N�X�|�[�g���Ă���낵���ł����H" & vbCrLf _
                & vbCrLf _
                & "�R���|�[�l���g���F %s" & vbCrLf _
                & "�G�N�X�|�[�g��F %s"
        
        ret.Add getMsgKey(msg.NONE_EXPORTPATH, lang.English), _
                "The following component will be not exported because the export path is not set." & vbCrLf _
                & vbCrLf _
                & "Component: %s"
        ret.Add getMsgKey(msg.NONE_EXPORTPATH, lang.Japanese), _
                "�ȉ��̃R���|�[�l���g�̓G�N�X�|�[�g�悪�ݒ肳��Ă��Ȃ����߃G�N�X�|�[�g����܂���B" & vbCrLf _
                & vbCrLf _
                & "�R���|�[�l���g���F %s"
        
        ret.Add getMsgKey(msg.NONE_CONFIG, lang.English), _
                "The config file is not found in ""%s""."
        ret.Add getMsgKey(msg.NONE_CONFIG, lang.Japanese), _
                """%s""�ɐݒ�t�@�C����������܂���B"
        
        ret.Add getMsgKey(msg.NONE_ROOTPATH, lang.English), _
                "Importing component and creating menu will be not executed for %s." & vbCrLf _
                & "Because the following directory is not found." & vbCrLf _
                & vbCrLf _
                & "Directory: %s"
        ret.Add getMsgKey(msg.NONE_ROOTPATH, lang.Japanese), _
                "�ȉ��̃t�H���_��������Ȃ����߁A" & vbCrLf _
                & "%s�̃R���|�[�l���g�̃C���|�[�g�y�у��j���[�����͎��s����܂���B" & vbCrLf _
                & vbCrLf _
                & "�t�H���_�F %s"
        
        ret.Add getMsgKey(msg.INFO_FOLDER, lang.English), _
                "Directory: %s" & vbCrLf
        ret.Add getMsgKey(msg.INFO_FOLDER, lang.Japanese), _
                "�t�H���_�F %s" & vbCrLf
        
        ret.Add getMsgKey(msg.INFO_FILE, lang.English), _
                "File: %s" & vbCrLf
        ret.Add getMsgKey(msg.INFO_FILE, lang.Japanese), _
                "�t�@�C���F %s" & vbCrLf
        
    End If
    Set getMsgManager = ret
    Exit Function

CATCH_ERR:
    Err.Raise Err.Number, "getMsgManager > " & Err.Source, Err.Description
End Function

Private Function getMsgKey(ByVal msgtype As msg, ByVal lang As lang) As String
    getMsgKey = msgtype & ":" & lang
End Function

