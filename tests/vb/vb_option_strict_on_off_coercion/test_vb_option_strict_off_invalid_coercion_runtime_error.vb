' vybe-test: vb/vb_option_strict_on_off_coercion/test_vb_option_strict_off_invalid_coercion_runtime_error
' origin: languages/vb/tests/vb/test_vb_option_strict_on_off_coercion.rs

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

Option Strict Off
Imports System

Module Program
    Sub Main()
        Dim invalidStr As String = "ABC"
        Try
            Dim num As Integer = invalidStr ' Runtime failure when coercion fails!
        Catch ex As InvalidCastException
            __Check(CStr("InvalidCastException Caught on Coercion"), "InvalidCastException Caught on Coercion")
        End Try
    End Sub
End Module
