' vybe-test: vb/vb_do_loop_while_until_variants/test_vb_do_while_complex_boolean_expression
' origin: languages/vb/tests/vb/test_vb_do_loop_while_until_variants.rs

Module Program
    Sub Main()
        Dim a = 0
        Dim b = 10
        Do While a < 5 AndAlso b > 5
            a += 1
            b -= 1
        Loop
        Console.WriteLine(a & "|" & b)
    End Sub
End Module
