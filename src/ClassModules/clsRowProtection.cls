VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRowProtection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' �� �s�}���^�폜�Ď��p�N���X


' �� �萔�E�ϐ���` {
Private Const DefaultConfigSheetName = "RowProtectionConfig" ' �f�t�H���g�ݒ�p�V�[�g��

Private ConfigWorkSheet As Worksheet ' �ݒ�p�V�[�g
Private WatchWorksheet As Worksheet ' �Ď��V�[�g
Private RegExpRowNumber As Object ' �s�ԍ��擾�p���K�\���I�u�W�F�N�g

' }


' �� �v���V�[�W�� {

Public Sub Init(Optional ConfigSheetName As String, Optional TargetSheet As Worksheet)
    If ConfigSheetName = "" Then ConfigSheetName = DefaultConfigSheetName
    Set ConfigWorkSheet = Worksheets(ConfigSheetName)
    
    Set WatchWorksheet = IIf(TargetSheet Is Nothing, ActiveSheet, TargetSheet)
    
    Set RegExpRowNumber = CreateObject("VBScript.RegExp")
    With RegExpRowNumber
        .Pattern = "\d+$"
        .IgnoreCase = False
        .Global = True
    End With
End Sub

Public Function ProtectRowInsert(ByRef TargetRange As Range) As Boolean
    ProtectRowInsert = False

    Dim UnmatchWidth As Boolean: UnmatchWidth = False
    Dim WatchSheetWidth As Long: WatchSheetWidth = WatchWorksheet.Columns.Count
    Dim TargetRangeWidth As Long: TargetRangeWidth = TargetRange.Columns.Count
    Dim TargetRangeHeight As Long: TargetRangeHeight = TargetRange.Rows.Count
    Dim TableColumn As Long: TableColumn = 0
    Dim TableWidth As Long: TableWidth = 0
    
    ' �Ώ�Range���V�[�g�������ς��ł��邩����
    If TargetRangeWidth <> WatchSheetWidth Or TargetRange.CountLarge Mod WatchSheetWidth <> 0 Then
        UnmatchWidth = True
        
        ' �Ώ�Range���e�[�u���������ς��ł��邩����
        
        On Error Resume Next
        With TargetRange.ListObject
            ' �e�[�u���̃J�����ʒu�ƕ��擾
            TableColumn = .Range.Column
            TableWidth = .ListColumns.Count
        End With
        
        If TableColumn = 0 Then
            ' �e�[�u���w�b�_(1�s��)�̒��O�ɑ}�����悤�Ƃ����P�[�X�ɑΉ�
            With TargetRange.Offset(TargetRangeHeight, 0).ListObject
                ' �e�[�u���̃J�����ʒu�ƕ��擾
                TableColumn = .Range.Column
                TableWidth = .ListColumns.Count
            End With
        End If
        On Error GoTo 0
        If TableColumn = TargetRange.Column And TableWidth = TargetRangeWidth Then UnmatchWidth = False
        
        If UnmatchWidth Then Exit Function
    End If
    
    Dim WatchSheetName As String: WatchSheetName = WatchWorksheet.name
    Dim SearchStrings(2) As String
    Dim SearchStringIndex As Long
    Dim FirstCell As Range
    Dim CurrentCell As Range
    Dim LockRowNumber As Long
    Dim AboveRowNumber As Long
        
    With ConfigWorkSheet.Range("A:A")
        ' ��������������쐬
        SearchStrings(1) = "=" & WatchSheetName & "!"
        SearchStrings(2) = "='" & WatchSheetName & "'!"
        ' �����V�[�g�̃Z���Q�Ǝ��ɂ́A�V�[�g����'�`'�ň͂܂�Ă���ꍇ�ƈ͂܂�Ă��Ȃ��ꍇ�Ƃ�����
        
        For SearchStringIndex = 1 To 2
            ' �ݒ�V�[�g�̃��b�N�w��Z���iA��j�̂����A�����ɑΏۃV�[�g�����܂ނ��̂̂݌���
            Set FirstCell = .Find(What:=SearchStrings(SearchStringIndex), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows)
            
            If Not FirstCell Is Nothing Then
                Set CurrentCell = FirstCell
                
                Do
                    LockRowNumber = GetRowNumber(CurrentCell) ' ���b�N�w��Z���s�ԍ�
                    If LockRowNumber < 0 Then
                        GoTo UNDO_CHANGE
                    End If
                    
                    If 0 < LockRowNumber Then
                        AboveRowNumber = GetRowNumber(CurrentCell.Offset(0, 1)) ' ����w��Z���s�ԍ�
                        If AboveRowNumber < 0 Then
                            ' ����w��Z�����Q�ƃG���[(#REF!)�ɂȂ��Ă���ꍇ�A���b�N�w��Z����1�s�ڂ̏ꍇ�̂݋��e
                            If LockRowNumber <> 1 Then GoTo UNDO_CHANGE
                        ElseIf LockRowNumber <> AboveRowNumber + 1 Then
                            ' ����w��Z�������b�N�w��Z���̒���ɖ����ꍇ�ɂ�Undo
                            GoTo UNDO_CHANGE
                        End If
                    End If
                    
                    Set CurrentCell = .FindNext(CurrentCell)
                    
                    If CurrentCell.Address = FirstCell.Address Then Exit Do
                Loop
            End If
        Next
    End With
    
    Exit Function
    
UNDO_CHANGE:
    Application.EnableEvents = False
    Application.Undo
    Application.EnableEvents = True
    
    ProtectRowInsert = True ' Undo ���ꂽ�ꍇ�ATrue ��Ԃ�
End Function

Private Function GetRowNumber(ByRef InfoCell As Range) As Long
    Dim Formula As String: Formula = InfoCell.Formula
    
    If Formula = "" Then
        GetRowNumber = 0
        Exit Function
    End If
    
    If Formula Like "*[#]REF!" Then
        GetRowNumber = -1
        Exit Function
    End If
            
    Dim Matches As Object: Set Matches = RegExpRowNumber.Execute(Formula)
    If 0 < Matches.Count Then
        GetRowNumber = Matches(0)
    Else
        GetRowNumber = -1
    End If
End Function

' }
