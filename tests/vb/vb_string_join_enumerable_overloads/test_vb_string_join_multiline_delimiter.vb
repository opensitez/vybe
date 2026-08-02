' vybe-test: vb/vb_string_join_enumerable_overloads/test_vb_string_join_multiline_delimiter
' origin: languages/vb/tests/vb/test_vb_string_join_enumerable_overloads.rs

Module Program
    Sub Main()
        Dim lines = {"Header", "Body", "Footer"}
        Console.WriteLine(String.Join(vbCrLf, lines))
    End Sub
End Module
