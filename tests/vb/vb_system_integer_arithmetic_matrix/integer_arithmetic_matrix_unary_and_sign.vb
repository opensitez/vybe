' vybe-test: vb/vb_system_integer_arithmetic_matrix/integer_arithmetic_matrix_unary_and_sign
' origin: languages/vb/tests/vb/test_vb_system_integer_arithmetic_matrix.rs

Module M
    Sub Main()
        Dim values() As Integer = {0, 1, -1, 2, -3, 9}
        Dim allGood As Boolean = True

        For Each x In values
            If ((+x) <> x) Then allGood = False
            If ((-x) <> (0 - x)) Then allGood = False
            If (Math.Sign(x) > 0 AndAlso x <= 0) Then allGood = False
        Next

        Console.WriteLine(allGood)
    End Sub
End Module
