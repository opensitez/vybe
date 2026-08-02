' vybe-test: vb/vb_system_array_2d_matrix/array_2d_matrix_fill_with_nested_loops_then_scan
' origin: languages/vb/tests/vb/test_vb_system_array_2d_matrix.rs

Imports System

Module M
    Sub Main()
        Dim m(0 To 3, 0 To 1) As Integer
        Dim seed As Integer = 1
        For r As Integer = m.GetLowerBound(0) To m.GetUpperBound(0)
            For c As Integer = m.GetLowerBound(1) To m.GetUpperBound(1)
                m(r, c) = seed
                seed += 1
            Next
        Next

        Dim total As Integer = 0
        For r As Integer = m.GetLowerBound(0) To m.GetUpperBound(0)
            For c As Integer = m.GetLowerBound(1) To m.GetUpperBound(1)
                total += m(r, c)
            Next
        Next

        Console.WriteLine(total)
        Console.WriteLine(seed)
    End Sub
End Module
