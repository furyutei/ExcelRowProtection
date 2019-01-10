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

' ■ 行挿入／削除監視用クラス


' ◆ 定数・変数定義 {
Private Const DefaultConfigSheetName = "RowProtectionConfig" ' デフォルト設定用シート名

Private ConfigWorkSheet As Worksheet ' 設定用シート
Private WatchWorksheet As Worksheet ' 監視シート
Private RegExpRowNumber As Object ' 行番号取得用正規表現オブジェクト

' }


' ◆ プロシージャ {

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
    
    ' 対象Rangeがシート幅いっぱいであるか判定
    If TargetRangeWidth <> WatchSheetWidth Or TargetRange.CountLarge Mod WatchSheetWidth <> 0 Then
        UnmatchWidth = True
        
        ' 対象Rangeがテーブル幅いっぱいであるか判定
        
        On Error Resume Next
        With TargetRange.ListObject
            ' テーブルのカラム位置と幅取得
            TableColumn = .Range.Column
            TableWidth = .ListColumns.Count
        End With
        
        If TableColumn = 0 Then
            ' テーブルヘッダ(1行目)の直前に挿入しようとしたケースに対応
            With TargetRange.Offset(TargetRangeHeight, 0).ListObject
                ' テーブルのカラム位置と幅取得
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
        ' 数式検索文字列作成
        SearchStrings(1) = "=" & WatchSheetName & "!"
        SearchStrings(2) = "='" & WatchSheetName & "'!"
        ' ※他シートのセル参照時には、シート名が'～'で囲まれている場合と囲まれていない場合とがある
        
        For SearchStringIndex = 1 To 2
            ' 設定シートのロック指定セル（A列）のうち、数式に対象シート名を含むもののみ検索
            Set FirstCell = .Find(What:=SearchStrings(SearchStringIndex), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows)
            
            If Not FirstCell Is Nothing Then
                Set CurrentCell = FirstCell
                
                Do
                    LockRowNumber = GetRowNumber(CurrentCell) ' ロック指定セル行番号
                    If LockRowNumber < 0 Then
                        GoTo UNDO_CHANGE
                    End If
                    
                    If 0 < LockRowNumber Then
                        AboveRowNumber = GetRowNumber(CurrentCell.Offset(0, 1)) ' 直上指定セル行番号
                        If AboveRowNumber < 0 Then
                            ' 直上指定セルが参照エラー(#REF!)になっている場合、ロック指定セルが1行目の場合のみ許容
                            If LockRowNumber <> 1 Then GoTo UNDO_CHANGE
                        ElseIf LockRowNumber <> AboveRowNumber + 1 Then
                            ' 直上指定セルがロック指定セルの直上に無い場合にはUndo
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
    
    ProtectRowInsert = True ' Undo された場合、True を返す
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

