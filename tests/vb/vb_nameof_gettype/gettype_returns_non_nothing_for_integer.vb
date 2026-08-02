' vybe-test: vb/vb_nameof_gettype/gettype_returns_non_nothing_for_integer
' origin: languages/vb/tests/vb/test_vb_nameof_gettype.rs

Module M
    Sub Main()
        Dim t As Object = GetType(Integer)
        If IsNothing(t) Then
            Console.WriteLine("missing")
        Else
            Console.WriteLine("present")
        End If
    End Sub
End Module
