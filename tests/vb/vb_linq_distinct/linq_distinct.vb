' vybe-test: vb/vb_linq_distinct/linq_distinct
' origin: languages/vb/tests/vb/test_vb_linq_distinct.rs

Imports System.Linq

Module M
    Sub Main()
        Dim numbers = {1, 2, 2, 3, 3, 3}
        
        Dim query = (From n In numbers Select n).Distinct()
        
        For Each n In query
            Console.WriteLine(n)
        Next
    End Sub
End Module
