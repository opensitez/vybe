' vybe-test: vb/vb_linq_group_by_into/linq_group_by_into
' origin: languages/vb/tests/vb/test_vb_linq_group_by_into.rs

Imports System.Linq

Module M
    Sub Main()
        Dim words = {"apple", "banana", "apricot", "cherry"}
        
        Dim query = From w In words
                    Group By Key = w.Substring(0, 1) Into Group, Count()
                    Order By Key
                    
        For Each g In query
            Console.WriteLine(g.Key & g.Count)
        Next
    End Sub
End Module
