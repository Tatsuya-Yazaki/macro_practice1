# macro_practice1
Sub AddCal()
'
' AddCal Macro
'

'
    Sheets("マスター").Select
    Sheets("マスター").Copy After:=Sheets(2)
    ActiveSheet.Name = Range("マスター!$D$2").Value & Range("マスター!$C$2").Value & Range("マスター!$D$3").Value & Range("マスター!$C$3").Value
    Range("C2:I29").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B5").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("B4").Select
End Sub
