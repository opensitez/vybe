' vybe-test: vb/vb_system_integer_arithmetic_matrix/integer_arithmetic_matrix_bitwise_compositions
' origin: languages/vb/tests/vb/test_vb_system_integer_arithmetic_matrix.rs

Module M
    Sub Main()
        Dim values() As Integer = {0, 1, 2, 3, 5, 8, 13}
        Dim allGood As Boolean = True

        For Each a In values
            If ((a << 1) >> 1 <> a) Then allGood = False
            If ((a >> 1) <= a) = False Then allGood = False
            If ((a Xor a) <> 0) Then allGood = False
            If ((a Or 0) <> a) Then allGood = False
            If ((a And 0) <> 0) Then allGood = False
        Next

        Console.WriteLine(allGood)
    End Sub
End Module
