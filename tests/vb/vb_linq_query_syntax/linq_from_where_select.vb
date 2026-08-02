' vybe-test: vb/vb_linq_query_syntax/linq_from_where_select
' origin: languages/vb/tests/vb/test_vb_linq_query_syntax.rs

Module M
    Sub Main()
        Dim numbers As Integer() = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}
        
        Dim evens = From n In numbers
                    Where n Mod 2 = 0
                    Select n
                    
        For Each e In evens
            Console.WriteLine(e)
        Next
    End Sub
End Module
