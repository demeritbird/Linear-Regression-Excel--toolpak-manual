Attribute VB_Name = "datasetcleaning"
Option Explicit

Sub clear_all_rows_loop()
'Random code here
    Dim idx As Byte

   Application.ScreenUpdating = False
    
   For idx = 1 To 19 '19 being the total columns in our current DataSet.
        On Error GoTo SkipNext
      DataSet.Columns(idx).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
SkipNext:
   Next idx

    Application.ScreenUpdating = True

End Sub
