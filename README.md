# Exam-TestCode
learn github

①シート複製コピー（Aを配列→フォーマットのコピー作成）

Dim VarArray As Variant

VarArray = Range("a1").CurrentRegion.Value

Dim VarData As Variant
Dim StrData As String

For Each VarData In VarArray

    Sheets("フォーマット").Copy Before:=Sheets("フォーマット")
	ActiveSheet.Name = VarData
    Range("C5").Value = VarData

Next VarData


②ハイパーリング

    'シート名の文字列を保持します
    Dim namesArray() As String
    ReDim namesArray(Sheets.Count)
    'シートの一覧を取得します
    For cnt = 1 To Sheets.Count
        namesArray(cnt) = Sheets(cnt).Name
    Next
    '一覧用のシートを追加します
    Worksheets.Add Before:=Worksheets(1)
    Set newWorkSheet = Worksheets(1)
    'ハイパーリンクの設定
    For cnt = 1 To UBound(namesArray)
        newWorkSheet.Hyperlinks.Add Anchor:=newWorkSheet.Cells(cnt, 1), Address:="", _
                         SubAddress:=namesArray(cnt) & "!A1", TextToDisplay:=namesArray(cnt)
    Next


③A1でカーソル初期化

    Dim s As Object
    Dim defaultSheet As Object
    Set defaultSheet = ActiveSheet
    For Each s In ActiveWorkbook.Sheets
        s.Activate
        ActiveSheet.Range("A1").Select
        ActiveWindow.Zoom = 100
    Next s
    defaultSheet.Activate
