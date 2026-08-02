' vybe-test: vb/vb_str_val_conversion_functions/test_vb_hex_function_integer_formatting
' origin: languages/vb/tests/vb/test_vb_str_val_conversion_functions.rs

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

Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        __Check(CStr(Hex(255) & "|" & Hex(16)), "FF|10")
    End Sub
End Module
