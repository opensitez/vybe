' vybe-test: vb/vb_builtins/foreach_loop
' origin: languages/vb/tests/vb/vb_builtins_test.rs

Module Program
    Sub Main()
        Dim items() As String = {"apple", "banana", "cherry"}
        For Each item As String In items
            Console.WriteLine(item)
        Next
    End Sub
End Module
