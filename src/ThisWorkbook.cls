VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const MYMENUNM As String = "VBAPorter"

Private Sub Workbook_AddinInstall()
    deleteMyMenu
    createMyMenu
End Sub

Private Sub Workbook_AddinUninstall()
    deleteMyMenu
End Sub

Private Sub Workbook_Open()
    main.initialize
End Sub


Private Sub createMyMenu()
    Dim bar As CommandBar
    Dim mymenu As CommandBarPopup
    Dim menubtn As CommandBarButton

    Set bar = Application.CommandBars("Worksheet Menu Bar")
    Set mymenu = bar.Controls.Add(Type:=MsoControlType.msoControlPopup)
    mymenu.Caption = MYMENUNM
    Set menubtn = mymenu.Controls.Add(Type:=MsoControlType.msoControlButton)
    menubtn.Caption = "全て更新"
    menubtn.OnAction = "main.updateAll"
End Sub

Private Sub deleteMyMenu()
    Dim bar As CommandBar
    Dim i As Integer

    For i = 1 To bar.Controls.Count
        If bar.Controls.Item(i).Caption = MYMENUNM Then
            bar.Controls.Item(i).Delete
            Exit For
        End If
    Next
End Sub


