' vybe-test: vb/vb_linq_orderby_desc_multi/linq_orderby_desc_multi
' origin: languages/vb/tests/vb/test_vb_linq_orderby_desc_multi.rs

Imports System.Linq

Class Item
    Public Id As Integer
    Public Val As Integer
End Class

Module M
    Sub Main()
        Dim items = {New Item With {.Id = 1, .Val = 10}, New Item With {.Id = 2, .Val = 10}, New Item With {.Id = 3, .Val = 5}}
        
        Dim query = From i In items
                    Order By i.Val Descending, i.Id Ascending
                    Select i.Id
                    
        For Each id In query
            Console.WriteLine(id)
        Next
    End Sub
End Module
