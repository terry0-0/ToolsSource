Attribute VB_Name = "Constant"
Option Explicit


'''�V�[�g����������
Function SheetSearchStr() As Range
    Set SheetSearchStr = ThisWorkbook.Sheets("�ݒ�").Cells(5, 3)
End Function

'''�����Z���ꊇ���͕�����
Function MultiInputString() As Range
    Set MultiInputString = ThisWorkbook.Sheets("�ݒ�").Cells(7, 3)
End Function

'''�����Z���ꊇ���͏㏑���ݒ�
Function MultiInputOverwrite() As Range
    Set MultiInputOverwrite = ThisWorkbook.Sheets("�ݒ�").Cells(8, 3)
End Function


'''�u�������u�b�N
Function ReplaceDicBook() As Range
    Set ReplaceDicBook = ThisWorkbook.Sheets("�ݒ�").Cells(11, 3)
End Function

'''�u�������V�[�g
Function ReplaceDicSht() As Range
    Set ReplaceDicSht = ThisWorkbook.Sheets("�ݒ�").Cells(12, 3)
End Function

'''�u�������J�n�s
Function ReplaceDicStartRow() As Range
    Set ReplaceDicStartRow = ThisWorkbook.Sheets("�ݒ�").Cells(13, 3)
End Function

'''�u������������
Function ReplaceDicSearchCol() As Range
    Set ReplaceDicSearchCol = ThisWorkbook.Sheets("�ݒ�").Cells(14, 3)
End Function

'''�u�������u����
Function ReplaceDicRepCol() As Range
    Set ReplaceDicRepCol = ThisWorkbook.Sheets("�ݒ�").Cells(15, 3)
End Function

'''�O���b�h�͈�
Function SelectionGriding() As Range
    Set SelectionGriding = ThisWorkbook.Sheets("�ݒ�").Cells(17, 3)
End Function

'''PCL���f�t���O
Function pclMark() As Range
    Set pclMark = ThisWorkbook.Sheets("�ݒ�").Cells(18, 3)
End Function

'''���K�\���ŃZ�����N���A����ݒ�
Function ClearByRegRepSetting() As Range
    Set ClearByRegRepSetting = ThisWorkbook.Sheets("�ݒ�").Cells(20, 3)
End Function
