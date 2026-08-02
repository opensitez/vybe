' vybe-test: vb/vb_linq_query_syntax/linq_order_by
' origin: languages/vb/tests/vb/test_vb_linq_query_syntax.rs

Module M
    Sub Main()
        Dim words As String() = {"apple", "cherry", "banana"}
        
        Dim sorted = From w In words
                     Order By w Descending
                     Select w
                     
        For Each w In sorted
            Console.WriteLine(w)
        Next
    End Sub
End Module
