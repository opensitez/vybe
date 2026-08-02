' vybe-test: vb/vb_option_strict_off_type_coercion/test_vb_option_strict_off_invalid_string_conversion_throws
' origin: languages/vb/tests/vb/test_vb_option_strict_off_type_coercion.rs

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
Option Strict Off

Module Program
    Sub Main()
        Dim badStr As Object = "NotANumber"
        Try
            Dim n As Integer = badStr
        Catch ex As InvalidCastException
            __Check(CStr("InvalidCastException Caught on String to Integer"), "InvalidCastException Caught on String to Integer")
        End Try
    End Sub
End Module
