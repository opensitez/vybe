' vybe-test: vb/vb_linq_let/linq_let_clause
' origin: languages/vb/tests/vb/test_vb_linq_let.rs

Module M
    Sub Main()
        Dim words As String() = {"apple", "banana", "cherry"}
        
        Dim query = From w In words
                    Let len = w.Length
                    Where len > 5
                    Select w & " is " & len.ToString()
                    
        For Each item In query
            Console.WriteLine(item)
        Next
    End Sub
End Module
