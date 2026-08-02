' vybe-test: vb/vb_query_select_anonymous/query_select_anonymous
' origin: languages/vb/tests/vb/test_vb_query_select_anonymous.rs

Imports System.Linq

Module M
    Sub Main()
        Dim ids = {1, 2}
        
        Dim query = From id In ids
                    Select New With {.Index = id, .Name = "Item" & id}
                    
        For Each item In query
            Console.WriteLine(item.Index & "-" & item.Name)
        Next
    End Sub
End Module
