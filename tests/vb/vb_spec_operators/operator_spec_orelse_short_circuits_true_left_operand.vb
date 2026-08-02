' vybe-test: vb/vb_spec_operators/operator_spec_orelse_short_circuits_true_left_operand
' origin: languages/vb/tests/vb/test_vb_spec_operators.rs

Module M
    Function Explode() As Boolean
        Console.WriteLine("boom")
        Return False
    End Function

    Sub Main()
        Console.WriteLine(True OrElse Explode())
    End Sub
End Module
