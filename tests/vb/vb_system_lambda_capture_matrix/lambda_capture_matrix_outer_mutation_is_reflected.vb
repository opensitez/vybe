' vybe-test: vb/vb_system_lambda_capture_matrix/lambda_capture_matrix_outer_mutation_is_reflected
' origin: languages/vb/tests/vb/test_vb_system_lambda_capture_matrix.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Imports System

Module M
    Sub Main()
        Dim factor As Integer = 2
        Dim fn As Func(Of Integer, Integer) = Function(v As Integer) v * factor

        factor = 5
        __Check(CStr(fn(3)), "15")
    End Sub
End Module
