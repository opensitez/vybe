' vybe-test: vb/vb_system_lambda_capture_matrix/lambda_capture_matrix_local_copy_via_for_loop_capture
' origin: languages/vb/tests/vb/test_vb_system_lambda_capture_matrix.rs

Imports System
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim funcs As New List(Of Func(Of Integer))()

        For i As Integer = 1 To 3
            Dim copied As Integer = i
            funcs.Add(Function() copied * 10)
        Next

        Dim a As Integer = funcs(0)()
        Dim b As Integer = funcs(1)()
        Dim c As Integer = funcs(2)()

        Console.WriteLine(a)
        Console.WriteLine(b)
        Console.WriteLine(c)
    End Sub
End Module
