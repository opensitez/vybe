' vybe-test: vb/vb_system_floating_point_matrix/floating_point_matrix_integer_conversion_roundtrip
' origin: languages/vb/tests/vb/test_vb_system_floating_point_matrix.rs

Module M
    Sub Main()
        Dim values() As Double = {1.2, -1.8, 9.99, 0.0}
        Dim allGood As Boolean = True

        For Each value In values
            Dim rounded As Double = Math.Round(value)
            Dim back As Double = CDbl(CInt(rounded))
            If back <> CInt(rounded) Then
                allGood = False
            End If
        Next

        Console.WriteLine(allGood)
    End Sub
End Module
