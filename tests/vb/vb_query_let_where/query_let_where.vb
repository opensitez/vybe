' vybe-test: vb/vb_query_let_where/query_let_where
' origin: languages/vb/tests/vb/test_vb_query_let_where.rs

Imports System.Linq

Module M
    Sub Main()
        Dim numbers = {1, 2, 3, 4, 5}
        
        Dim query = From n In numbers
                    Let doubled = n * 2
                    Where doubled > 5
                    Select doubled
                    
        For Each d In query
            Console.WriteLine(d)
        Next
    End Sub
End Module
