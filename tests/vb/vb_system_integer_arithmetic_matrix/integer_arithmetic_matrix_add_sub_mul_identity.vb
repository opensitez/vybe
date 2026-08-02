' vybe-test: vb/vb_system_integer_arithmetic_matrix/integer_arithmetic_matrix_add_sub_mul_identity
' origin: languages/vb/tests/vb/test_vb_system_integer_arithmetic_matrix.rs

Module M
    Sub Main()
        Dim values() As Integer = {-12, -3, 0, 1, 2, 5, 10, 17}

        Dim allGood As Boolean = True

        For Each a In values
            For Each b In values
                If (a + b - b <> a) Then
                    allGood = False
                End If
                If (a - b + b <> a) Then
                    allGood = False
                End If
                If (a * 1 <> a) Then
                    allGood = False
                End If
            Next
        Next

        Console.WriteLine(allGood)
    End Sub
End Module
