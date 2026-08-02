' vybe-test: vb/vb_advanced_linq_xml/linq_join_multiple_keys
' origin: languages/vb/tests/vb/test_vb_advanced_linq_xml.rs

Imports System.Linq

Class Item
    Public K1 As Integer
    Public K2 As Integer
    Public Val As String
End Class

Module M
    Sub Main()
        Dim arr1 = {New Item With {.K1 = 1, .K2 = 2, .Val = "A"}}
        Dim arr2 = {New Item With {.K1 = 1, .K2 = 2, .Val = "B"}}
        
        Dim query = From a In arr1
                    Join b In arr2 On a.K1 Equals b.K1 And a.K2 Equals b.K2
                    Select a.Val & b.Val
                    
        For Each res In query
            Console.WriteLine(res)
        Next
    End Sub
End Module
