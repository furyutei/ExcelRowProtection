VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ◆ ライン挿入／削除監視オブジェクト取得用プロシージャ {

Private RowProtectionObject As clsRowProtection

Private Function GetRowProtectionObject() As clsRowProtection
    If RowProtectionObject Is Nothing Then
        Set RowProtectionObject = New clsRowProtection
        Call RowProtectionObject.Init
        ' ※定義シート名を指定したい場合、Init(ConfigSheetName:="（シート名）")
    End If
    
    Set GetRowProtectionObject = RowProtectionObject
End Function

'}


' ◆ イベントプロシージャ {

Private Sub Worksheet_Activate()
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim IsUndone As Boolean: IsUndone = GetRowProtectionObject().ProtectRowInsert(Target)
End Sub

' }

