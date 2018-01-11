Imports Microsoft.Office.Interop.Excel

Namespace xlWorkSheet
    Friend Class Cells
        Friend Class SpecialCells
            Private xlCellTypeLastCell As XlCellType

            Public Sub New(xlCellTypeLastCell As XlCellType)
                Me.xlCellTypeLastCell = xlCellTypeLastCell
            End Sub
        End Class
    End Class
End Namespace
