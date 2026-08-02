' vybe-test: vb/vb_foreach_array_literal/foreach_array_literal
' origin: languages/vb/tests/vb/test_vb_foreach_array_literal.rs

Module M
    Sub Main()
        ' For Each with implicit array literal
        For Each x In {10, 20, 30}
            Console.WriteLine(x)
        Next
    End Sub
End Module
