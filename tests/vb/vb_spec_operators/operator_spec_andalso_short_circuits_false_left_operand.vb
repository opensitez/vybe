' vybe-test: vb/vb_spec_operators/operator_spec_andalso_short_circuits_false_left_operand
' origin: languages/vb/tests/vb/test_vb_spec_operators.rs

Module M
    Function Explode() As Boolean
        Console.WriteLine("boom")
        Return True
    End Function

    Sub Main()
        Console.WriteLine(False AndAlso Explode())
    End Sub
End Module
