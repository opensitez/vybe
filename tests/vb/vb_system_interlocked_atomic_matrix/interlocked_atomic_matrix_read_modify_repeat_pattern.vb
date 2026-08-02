' vybe-test: vb/vb_system_interlocked_atomic_matrix/interlocked_atomic_matrix_read_modify_repeat_pattern
' origin: languages/vb/tests/vb/test_vb_system_interlocked_atomic_matrix.rs

Imports System
Imports System.Threading

Module M
    Sub Main()
        Dim value As Integer = 2

        For i As Integer = 0 To 2
            Dim current As Integer = Interlocked.Increment(value)
        Next

        Dim before As Integer = value
        Dim compare As Integer = Interlocked.CompareExchange(value, 100, 5)

        Console.WriteLine(value)
        Console.WriteLine(compare)
    End Sub
End Module
