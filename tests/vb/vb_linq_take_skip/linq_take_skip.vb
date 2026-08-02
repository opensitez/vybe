' vybe-test: vb/vb_linq_take_skip/linq_take_skip
' origin: languages/vb/tests/vb/test_vb_linq_take_skip.rs

Imports System.Linq

Module M
    Sub Main()
        Dim numbers = {1, 2, 3, 4, 5}
        
        Dim query = From n In numbers
                    Skip 1
                    Take 2
                    Select n
                    
        For Each n In query
            Console.WriteLine(n)
        Next
    End Sub
End Module
