' vybe-test: vb/vb_string_composite_formatting_alignment/test_vb_string_format_throws_format_exception_missing_arg
' origin: languages/vb/tests/vb/test_vb_string_composite_formatting_alignment.rs

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

Module Program
    Sub Main()
        Try
            String.Format("{0} {1}", "OnlyOne")
        Catch ex As FormatException
            __Check(CStr("FormatException Missing Argument Caught"), "FormatException Missing Argument Caught")
        End Try
    End Sub
End Module
