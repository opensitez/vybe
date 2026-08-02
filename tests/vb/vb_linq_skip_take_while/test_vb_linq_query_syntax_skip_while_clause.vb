' vybe-test: vb/vb_linq_skip_take_while/test_vb_linq_query_syntax_skip_while_clause
' origin: languages/vb/tests/vb/test_vb_linq_skip_take_while.rs

Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"apple", "apricot", "banana", "cherry"}
        Dim query = From w In words Skip While w.StartsWith("a")
        Console.WriteLine(String.Join(",", query))
    End Sub
End Module
