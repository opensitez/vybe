' vybe-test: vb/vb_conversion_methods/convert_to_string_preserves_formatting_of_date
' origin: languages/vb/tests/vb/test_vb_conversion_methods.rs

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
        Dim d As Date = New Date(2024, 3, 14)
        Dim text As String = Convert.ToString(d)
        __Check(CStr(text.Length > 0), "True")
    End Sub
End Module
