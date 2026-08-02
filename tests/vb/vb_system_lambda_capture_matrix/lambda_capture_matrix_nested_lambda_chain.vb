' vybe-test: vb/vb_system_lambda_capture_matrix/lambda_capture_matrix_nested_lambda_chain
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
        Dim scale As Integer = 3
        Dim build As Func(Of Integer, Func(Of Integer)) = _
            Function(base As Integer) Function(x As Integer) (base + x) * scale

        Dim f As Func(Of Integer) = build(4)
        __Check(CStr(f()), "21")
    End Sub
End Module
