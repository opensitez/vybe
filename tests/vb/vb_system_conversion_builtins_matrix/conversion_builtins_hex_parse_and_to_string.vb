' vybe-test: vb/vb_system_conversion_builtins_matrix/conversion_builtins_hex_parse_and_to_string
' origin: languages/vb/tests/vb/test_vb_system_conversion_builtins_matrix.rs

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
        Dim value As Integer = Convert.ToInt32("ff", 16)
        Dim hex As String = Convert.ToString(value, 16)

        __Check(CStr(value), "255")
        __Check(CStr(hex), "ff")
        __Check(CStr(hex = "ff"), "True")
    End Sub
End Module
