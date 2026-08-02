' vybe-test: vb/vb_system_integer_arithmetic_matrix/integer_arithmetic_matrix_division_mod_contracts
' origin: languages/vb/tests/vb/test_vb_system_integer_arithmetic_matrix.rs

Module M
    Sub Main()
        Dim numerators() As Integer = {-8, -3, 2, 5, 17}
        Dim denominators() As Integer = {-3, 1, 2, 4}

        Dim allGood As Boolean = True

        For Each n In numerators
            For Each d In denominators
                Dim q As Integer = n \ d
                Dim r As Integer = n Mod d
                If (q * d + r <> n) Then
                    allGood = False
                End If
            Next
        Next

        Console.WriteLine(allGood)
    End Sub
End Module
