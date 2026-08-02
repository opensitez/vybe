' vybe-test: vb/vb_linq_let_complex/linq_let_complex
' origin: languages/vb/tests/vb/test_vb_linq_let_complex.rs

Imports System.Linq

Module M
    Sub Main()
        Dim names = {"Alice", "Bob", "Charlie"}
        
        Dim query = From n In names
                    Let len = n.Length
                    Where len > 3
                    Select n, len
                    
        For Each item In query
            Console.WriteLine(item.n & item.len)
        Next
    End Sub
End Module
