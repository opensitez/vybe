' vybe-test: vb/vb_linq_take_while/linq_take_while
' origin: languages/vb/tests/vb/test_vb_linq_take_while.rs

Imports System.Linq

Module M
    Sub Main()
        Dim nums = {1, 2, 3, 4, 1, 2}
        
        Dim query = From n In nums
                    Take While n < 4
                    Select n
                    
        For Each n In query
            Console.WriteLine(n)
        Next
    End Sub
End Module
