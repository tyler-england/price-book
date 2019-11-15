Attribute VB_Name = "Module2"
Sub DeleteOption()

Dim wksht As Worksheet, rows As Integer, allOptions As Range

Set wksht = Sheet3

wksht.Range("G1").Formula = "=sum(c:c)"

If wksht.Range("G1").Value > 0 Then

    msg = "WARNING!" & vbCrLf & "Are you sure you want to delete an option?" & vbCrLf & vbCrLf & "This process will delete it permanently."
    ans = MsgBox(msg, vbYesNo, "Delete Option(s)?")
    
    If ans = vbYes Then
        
        wksht.Activate
        
    rows = InputBox("What row number is the option you want to delete?")
    
    'delete row
    'delete option in costbook
        'update formulae/formatting
    'delete option in autoQuote
        'update formulae/sheet1/formatting
    
    Else: Cancel = True
    End If

End If

wksht.Activate
wksht.Range("G1").Clear
wksht.Range("A1").Select

End Sub

Sub AddOption()

Dim rot As Integer, semi As Integer, auto As Integer, price As Long
Dim singleHeads As Integer, twinHeads As Integer, row As Integer
Dim drive As Integer, desig1 As String, desig2 As String
Dim descSh As String, descLg As String, output As String
Dim scalable As Boolean, preEx As String, newRow As Integer
Dim updQuoteWB As Workbook, updCostWB As Workbook


rot = 0
semi = 0
auto = 0
singleHeads = 0
twinHeads = 0
drive = 0
desig1 = ""
desig2 = ""
price = 100000
scalable = False

Sheet3.Activate

MsgBox "You will be prompted to provide information for a new Mateer machine option. If you would like to change something after you've entered it, delete the new option and create the new option again."

descSh = InputBox("What is the short description of this option? (This is the text which will appear in this workbook, or in the price book.)")

'check if the short description is unique
Range("G1").Value = descSh
Range("G2").Formula = "=MATCH(G1,B:B,FALSE)"
row = 0
On Error Resume Next
row = Range("G2").Value
Range("G1:G2").Clear
If row > 0 Then
    MsgBox "There already exists an option by that name. Please enter a unique description."
    Exit Sub
End If

On Error GoTo errhandler

descLg = InputBox("What is the full description of this option? (This is the text which will appear as a line in the formal quote)")

msg = "Is this new option relevant to all Mateer filler models?"
ans = MsgBox(msg, vbYesNo, "Models")

If ans = vbNo Then
    msg = "Does this new option apply to rotary fillers?"
    ans = MsgBox(msg, vbYesNo, "Rotaries?")
    
    If ans = vbYes Then
        rot = 1
    Else
        rot = 0
    End If
    
    msg = "Does this new option apply to semiautomatic fillers?"
    ans = MsgBox(msg, vbYesNo, "Semiautomatics?")
    
    If ans = vbYes Then
        semi = 1
    Else
        semi = 0
    End If
    
    msg = "Does this new option apply to automatic fillers?"
    ans = MsgBox(msg, vbYesNo, "Automatics?")
    
    If ans = vbYes Then
        auto = 1
    Else
        auto = 0
    End If
    
Else
    rot = 1
    semi = 1
    auto = 1

End If

If rot = 1 And semi = 0 And auto = 0 Then
    desig1 = "Rotaries"
ElseIf rot = 0 And semi = 1 And auto = 1 Then
    desig1 = "Non-rotaries"
ElseIf rot = 0 And semi = 1 And auto = 0 Then
    desig1 = "Semiautomatic"
ElseIf rot = 0 And semi = 0 And auto = 1 Then
    desig1 = "Automatic"
Else
    desig1 = "All machines"
End If

msg = "Does this new option depend on whether the machine features single or twin fill heads?"
ans = MsgBox(msg, vbYesNo, "Heads?")

If ans = vbYes Then
    msg = "Does this option apply to twin heads? (2800, 2900, 6700)"
    ans = MsgBox(msg, vbYesNo, "Twin heads?")
    
    If ans = vbYes Then
        twinHeads = 1
        desig2 = "Twin Head"
    End If
    
    msg = "Does this option apply to single heads? (1100, 1200, 1800, 1900, automatics)"
    ans = MsgBox(msg, vbYesNo, "Single heads?")
    
    If ans = vbYes Then
        singleHeads = 1
        desig2 = "Single Head"
    End If
    
    If singleHeads = 1 And twinHeads = 1 Then
        MsgBox "It doesn't seem to matter what type of head the machine features. This option will be considered relevant to all types of heads."
        desig2 = ""
    End If

Else
    singleHeads = 1
    twinHeads = 1
    
End If

If rot + auto > 0 Then
    'question does not apply to semiautomatics
    msg = "Does this option scale with the number of columns? (e.g. it would cost twice as much for a 4900 than for a 3900)"
    ans = MsgBox(msg, vbYesNo, "Scalable?")
    
    If ans = vbYes Then
        scalable = True
    End If
End If

price = InputBox("What is the price for this option? (Enter number only)")

If singleHeads + twinHeads = 2 Then

    output = desig1
    
Else
    output = desig2
End If

Range("G1").Value = output
Range("G2").Formula = "=MATCH(G1,A:A,FALSE)"
row = Range("G2").Value

Range("B" & row).Select
Selection.End(xlDown).Select
row = ActiveCell.row

Do While Range("B" & row).Borders.Item(xlEdgeTop).LineStyle = xlLineStyleNone
    row = row - 1
Loop

preEx = Range("B" & row - 1).Value

Set updQuoteWB = Workbooks.Open(Filename:="K:\EnglandT\MATEER\QUOTES\Quote_Auto.xlsm", UpdateLinks:=True)
Set updCostWB = Workbooks.Open(Filename:="K:\EnglandT\MATEER\QUOTES\CostBook_Mateer.xlsm", UpdateLinks:=True)
ThisWorkbook.Activate

rows(row).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Range("B" & row & ":E" & row).Select
Selection.Borders.LineStyle = xlContinuous

Range("B" & row).Value = descSh
Range("C" & row).Value = price
If scalable = True Then
    Range("D" & row).Value = "Yes"
Else
    Range("D" & row).Value = "No"
End If
Range("E" & row - 1).Select
Selection.AutoFill Destination:=Range("E" & row - 1 & ":E" & row)
Range("F" & row).Value = descLg

'update AutoQuote workbook
newRow = NewOption(updQuoteWB, preEx)
updQuoteWB.Activate
Range("B" & newRow + 1 & ":N" & newRow + 1).Select
Selection.AutoFill Destination:=Range("B" & newRow & ":N" & newRow + 1)
Range("D" & newRow).Value = 1
Range("B" & newRow - 1).Select
Selection.AutoFill Destination:=Range("B" & newRow - 1 & ":B" & newRow + 1)
Range("E" & newRow).Value = Range("B" & newRow).Value
updQuoteWB.Activate
Sheet1.Activate
updQuoteWB.Worksheets("Options").Visible = False

updQuoteWB.Save
Application.DisplayAlerts = False
With updQuoteWB
    SetAttr .FullName, vbReadOnly
    .ChangeFileAccess xlReadOnly
End With
Application.DisplayAlerts = True
updQuoteWB.Close


ThisWorkbook.Activate
Range("G1:G2").Clear
Range("B" & row).Select

'Update costbook
newRow = NewOption(updCostWB, preEx)
updCostWB.Activate
Range("B" & newRow).Select
ActiveCell.Formula = "=[PriceBook_Mateer.xlsm]Options!B" & row

Exit Sub

errhandler: MsgBox "Unspecified error"
End Sub

Function NewOption(wkbk As Workbook, preEx As String) As Integer
Dim wksht As Worksheet, newRow As Integer

    wkbk.Activate

    If wkbk.ReadOnly Then
            
        With wkbk
            SetAttr .FullName, vbNormal
            .ChangeFileAccess xlReadWrite
    
            Application.DisplayAlerts = False
            .Save
            Application.DisplayAlerts = True
        End With
    
    End If
    
    Set wksht = Worksheets("Options")
    
    wksht.Activate
    wksht.Unprotect
    Range("O2").Value = preEx
    Range("O3").Formula = "=MATCH(O2,B:B,FALSE)"
    newRow = Range("O3").Value + 1
    Range("O2:O3").Clear
    rows(newRow).Insert
    NewOption = newRow

End Function
