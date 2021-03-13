Sub NewItem()
Dim rowEnd As Integer

    rowEnd = Cells(Rows.Count, "B").End(xlUp).Row
    

    Cells(rowEnd + 1, 2).Value = Cells(3, 2).Value
    Cells(rowEnd + 1, 3).Value = Cells(3, 3).Value
    Cells(rowEnd + 1, 4).Value = Cells(3, 4).Value
    Cells(rowEnd + 1, 5).Value = Cells(3, 5).Value
    Cells(rowEnd + 1, 6).Value = Cells(3, 6).Value
    
    Cells(rowEnd + 1, 1).Value = (rowEnd - 2) & "."
    Rows(3).ClearContents

    Range("B3").Activate
End Sub
Sub InputClear()
'-------------------Deletes all Inventory Data---------------------
Dim rowEnd As Integer
    rowEnd = Cells(Rows.Count, "B").End(xlUp).Row
    For i = 3 To rowEnd
        Rows(i).ClearContents
    Next
End Sub
Sub Undo()
'-------------------Getting Number of Last Row (Most Recent Addition) And Removing It---------
Dim rowEnd As Integer
    rowEnd = Cells(Rows.Count, "B").End(xlUp).Row
    Rows(rowEnd).Clear
End Sub
Sub InventorySheet()
    Range("B3").Activate
    
    Dim today As Date
    
    Dim f As Range
    Dim k As Range
    Dim e As Range
    
    Dim additem As Excel.Button
    Dim clearline As Excel.Button
    Dim Undo As Excel.Button
    
    today = Date
    Application.ScreenUpdating = False
    ActiveSheet.Buttons.Delete

'-------------Title information----------------------------
    Rows(1).RowHeight = 30
    Range("A1:C1").Merge
    Range("A1:C1").Value = "Inventory " & today
    Range("A1:C1").Font.Size = 20
    Range("A1:C1").HorizontalAlignment = xlCenter
    Range("A1:C1").VerticalAlignment = xlCenter
    
    Range("A1:F1").Interior.Color = RGB(186, 217, 207)
    Range("A1:F1").BorderAround LineStyle:=xlContinuous, ColorIndex:=1

'-------------Set Column Headers and Default Widths------------
    Cells(2, 1).Value = "#"
    Cells(2, 2).Value = "Item"
    Cells(2, 3).Value = "Serial Number"
    Cells(2, 4).Value = "Location"
    Cells(2, 5).Value = "Quantity"
    Cells(2, 6).Value = "Notes"
    

    
    Columns("A").ColumnWidth = 6
    Columns("B:C").ColumnWidth = 30
    Columns("D").ColumnWidth = 20
    Columns("E").ColumnWidth = 12
    Columns("F").ColumnWidth = 50
    


'-------------Borders and Color for Headers/Insert Cells-----------------
    Range("A3:F3").Interior.Color = vbYellow
    Range("A3:F3").BorderAround LineStyle:=xlContinuous, ColorIndex:=1
    Range("A3:F3").Borders(xlInsideVertical).LineStyle = xlContinuous
    
    Range("A2:F2").BorderAround LineStyle:=xlContinuous, ColorIndex:=1
    Range("A2:F2").Borders(xlInsideVertical).LineStyle = xlContinuous
    Range("A2:F2").HorizontalAlignment = xlCenter
    
    Rows(2).Font.Bold = True
    
'------------Creating Buttons
    Set f = Range("H3:I4")
    Set additem = ActiveSheet.Buttons.Add(f.Left, f.Top, f.Width, f.Height)
    additem.Caption = "Add Item"
    additem.OnAction = "NewItem"
    
    Set k = Range("H6:I7")
    Set clearline = ActiveSheet.Buttons.Add(k.Left, k.Top, k.Width, k.Height)
    clearline.Caption = "Clear Inventory"
    clearline.OnAction = "InputClear"
    
    Set e = Range("H9:I10")
    Set Undo = ActiveSheet.Buttons.Add(e.Left, e.Top, e.Width, e.Height)
    Undo.Caption = "Undo Last"
    Undo.OnAction = "Undo"

    
    
End Sub
