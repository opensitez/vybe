' vybe-test: vb/vb_system_lambda_capture_matrix/lambda_capture_matrix_closure_with_query_projection
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

Imports System.Linq

Module M
    Sub Main()
        Dim values() As Integer = {1, 2, 3, 4}
        Dim offset As Integer = 1

        Dim projected = values.Where(Function(v) (v + offset) Mod 2 = 0).Select(Function(v) v + offset)

        __Check(CStr(String.Join(",", projected)), "3,5")
    End Sub
End Module
