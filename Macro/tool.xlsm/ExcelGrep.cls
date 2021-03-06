VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelGrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Private MainSheet           As Worksheet

Private TargetPathCell      As Range

Private SearchPatternCell   As Range

Private ResultListFirstCell As Range

Private ShouldSearch        As Boolean

Private fso

Private REG                 As New VBScript_RegExp_55.RegExp

Private InvisibleExcel      As Excel.Application



Private Enum ResultColumns

    path = 1

    Book = 2

    Sheet = 3

    Name = 4

    Value = 5

End Enum



Private Sub Class_Initialize()

    Set MainSheet = ThisWorkbook.Sheets("ExcelGrep")

    Set TargetPathCell = MainSheet.Range("C3")

    Set SearchPatternCell = MainSheet.Range("C4")

    Set ResultListFirstCell = MainSheet.Range("C7")

    ShouldSearch = False

End Sub



Public Sub PickupFolderPath(Msg As String)

    Dim FolderPicker

    Dim Result       As String

    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)

    

    FolderPicker.Title = Msg

    If FolderPicker.Show Then Result = FolderPicker.SelectedItems(1)

    If Result = "" Then Exit Sub

    

    TargetPathCell.Value = Result

End Sub



Public Sub ExecSearch(Optional IgnoreCase As Boolean = True)

    Dim SpecifiedFolder
    Set SpecifiedFolder = GetSpecifiedFolder()

    If SpecifiedFolder Is Nothing Then

        MsgBox "検索対象フォルダが見つかりませんでした。", vbInformation, ThisWorkbook.Name

        Exit Sub

    End If

    If Trim(SearchPatternCell.Value) = "" Then

        MsgBox "検索文字列を入力してください。", vbInformation, ThisWorkbook.Name

        Exit Sub

    End If

    

    If ShouldSearch = True Then

        MsgBox "別の検索処理がすでに実行中です。", vbInformation, ThisWorkbook.Name

        Exit Sub

    End If

    

    REG.Global = True

    REG.IgnoreCase = IgnoreCase

    REG.pattern = SearchPatternCell.Value

    

    ShouldSearch = True

    Call ClearResultList

    Call SearchFolder(SpecifiedFolder)

    

    If Not InvisibleExcel Is Nothing Then

        InvisibleExcel.Quit

        Set InvisibleExcel = Nothing

    End If

    

    ShouldSearch = False

    Call DisplayStatus("")

    MsgBox "検索が終了しました。", vbInformation, ThisWorkbook.Name

End Sub



Public Sub Interrupt()

    If ShouldSearch = False Then Exit Sub

    If MsgBox("検索を中止してもよろしいですか?", vbYesNo Or vbQuestion, ThisWorkbook.Name) <> vbYes Then Exit Sub

    

    ShouldSearch = False

End Sub



Public Sub ClearResultList()

    ResultListFirstCell.CurrentRegion.Offset(2).Delete

End Sub



Private Function GetSpecifiedFolder()

    On Error Resume Next

    Dim SpecifiedFolder
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set SpecifiedFolder = fso.GetFolder(TargetPathCell.Value)

    Set GetSpecifiedFolder = SpecifiedFolder

End Function



Private Sub SearchFolder(objFolder)

    If ShouldSearch = False Then Exit Sub

    

    Dim objFile

    Dim SubFolder

    

    Call DisplayStatus(objFolder.path)

    

    For Each objFile In objFolder.Files

        Select Case fso.GetExtensionName(objFile.path)

            Case "xls", "xlsx", "xlsm"

                Call SearchBook(objFile)

        End Select

    Next

    

    For Each SubFolder In objFolder.SubFolders

        Call SearchFolder(SubFolder) 'recursive call

    Next

End Sub



Private Sub SearchBook(objFile)

    If ShouldSearch = False Then Exit Sub

    

    On Error Resume Next

    If InvisibleExcel Is Nothing Then

        Set InvisibleExcel = New Excel.Application

        InvisibleExcel.Visible = False

        InvisibleExcel.ScreenUpdating = False

    End If

    

    Dim Book As Workbook: Set Book = InvisibleExcel.Workbooks.Open(Filename:=objFile.path, readonly:=True)

    If Book Is Nothing Then

        MsgBox objFile.path & vbCrLf & " が開けませんでした。"

        Exit Sub

    End If

    On Error GoTo 0

    

    Dim Sheet      As Worksheet

    For Each Sheet In Book.Worksheets

        Call SearchSheet(Sheet)

    Next

    

    Call Book.Close(SaveChanges:=False)

End Sub



Private Sub SearchSheet(Sheet As Worksheet)

    If ShouldSearch = False Then Exit Sub

    

    Dim TargetRange As Range

    Dim Cell        As Range



    'Search Cells

    Set TargetRange = Sheet.UsedRange.Cells

    For Each Cell In TargetRange

        DoEvents

        Call DisplayStatus(Sheet.Parent.FullName)

        

        If Cell.Value <> "" Then

            If REG.test(Cell.Value) Then

                Call ProcessCell(Cell)

            End If

        End If

    Next

    

    'Search Shapes

    Dim objShape As Shape

    For Each objShape In Sheet.Shapes

        DoEvents

        Call DisplayStatus(Sheet.Parent.FullName)

        

        If HasTextFrameCharactersText(objShape) Then

            If REG.test(objShape.TextFrame.Characters.Text) Then

                Call ProcessShape(objShape)

            End If

        End If

    Next

End Sub



Private Function HasTextFrameCharactersText(objShape As Shape) As Boolean

    On Error Resume Next

    Dim Text As String

    Text = objShape.TextFrame.Characters.Text

    HasTextFrameCharactersText = (Text <> "")

End Function



Private Sub ProcessCell(Cell As Range)

    Call SetNewRowData(FoundSheet:=Cell.Parent, Name:=Cell.Address, Value:=Cell.Value)

End Sub



Private Sub ProcessShape(objShape As Shape)

    Call SetNewRowData(FoundSheet:=objShape.Parent, Name:=objShape.Name, Value:=objShape.TextFrame.Characters.Text)

End Sub



Private Sub SetNewRowData(FoundSheet As Worksheet, Name As String, Value As String)

    Dim Row      As Range: Set Row = GetNewRow()

    Dim PathCell As Range: Set PathCell = Row.Cells(ResultColumns.path)

    Dim BookCell As Range: Set BookCell = Row.Cells(ResultColumns.Book)

    

    'パス

    PathCell.Value = FoundSheet.Parent.FullName

    PathCell.WrapText = False

    Call MainSheet.Hyperlinks.Add(Anchor:=PathCell, Address:=PathCell.Value)

    'ブック

    BookCell.Value = FoundSheet.Parent.Name

    Call MainSheet.Hyperlinks.Add(Anchor:=BookCell, Address:=PathCell.Value)

    'シート

    Row.Cells(ResultColumns.Sheet).Value = FoundSheet.Name

    '名前

    Row.Cells(ResultColumns.Name).Value = Name

    '値

    Row.Cells(ResultColumns.Value).Value = Value

    Row.Cells(ResultColumns.Value).WrapText = False

    '--- 罫線 ---

    Row.Borders.LineStyle = xlContinuous

End Sub



Private Function GetNewRow() As Range

    Dim ListRange   As Range: Set ListRange = GetResultListRange

    Dim NewRowIndex As Long: NewRowIndex = ListRange.Rows.Count + 1

    Set GetNewRow = ListRange.Rows(NewRowIndex)

End Function



Private Function GetResultListRange() As Range

    Dim ListRowsCount As Long: ListRowsCount = ResultListFirstCell.CurrentRegion.Rows.Count - 1

    Set GetResultListRange = ResultListFirstCell.CurrentRegion.Offset(1).Resize(ListRowsCount)

End Function



Private Sub DisplayStatus(Msg As String)

    Const DotsLen  As Integer = 5

    Dim Dots       As String: Dots = String((Math.Rnd * DotsLen), ".")

    Dim Loading    As String: Loading = Left(Dots & String(DotsLen, " "), DotsLen)

    Dim strDisplay As String: strDisplay = "検索中" & Loading & " " & Msg

    Application.StatusBar = IIf(Msg = "", "", strDisplay)

    DoEvents

End Sub



Private Function IsIncludedInListRange(Target As Range) As Boolean

    Dim ResultRange As Range

    Set ResultRange = Application.Intersect(GetResultListRange, Target)

    IsIncludedInListRange = Not ResultRange Is Nothing

End Function



Public Sub FollowHyperlink(Target As Hyperlink)

    If Not IsIncludedInListRange(Target.Range) Then Exit Sub

    

    Dim SourceCell  As Range:     Set SourceCell = Target.Range

    Dim SourceSheet As Worksheet: Set SourceSheet = SourceCell.Parent

    Dim SourceRow   As Range:     Set SourceRow = SourceSheet.Range(SourceCell.End(xlToRight).End(xlToLeft), SourceCell.End(xlToRight))

    Dim SheetName   As String:        SheetName = SourceRow.Cells(ResultColumns.Sheet)

    Dim ObjectName  As String:        ObjectName = SourceRow.Cells(ResultColumns.Name)

    Dim Book        As Workbook:  Set Book = ActiveSheet.Parent

    

    On Error Resume Next

    Dim IsRangeObject As Boolean: IsRangeObject = (ObjectName Like "$*")

    Dim DistSheet     As Worksheet: Set DistSheet = Book.Sheets(SheetName)

    DistSheet.Activate

    If IsRangeObject Then

        DistSheet.Range(ObjectName).Activate

    Else

        DistSheet.Shapes(ObjectName).Select

    End If

End Sub
