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

' �� ���C���}���^�폜�Ď��I�u�W�F�N�g�擾�p�v���V�[�W�� {

Private RowProtectionObject As clsRowProtection

Private Function GetRowProtectionObject() As clsRowProtection
    If RowProtectionObject Is Nothing Then
        Set RowProtectionObject = New clsRowProtection
        Call RowProtectionObject.Init
        ' ����`�V�[�g�����w�肵�����ꍇ�AInit(ConfigSheetName:="�i�V�[�g���j")
    End If
    
    Set GetRowProtectionObject = RowProtectionObject
End Function

'}


' �� �C�x���g�v���V�[�W�� {

Private Sub Worksheet_Activate()
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim IsUndone As Boolean: IsUndone = GetRowProtectionObject().ProtectRowInsert(Target)
End Sub

' }
