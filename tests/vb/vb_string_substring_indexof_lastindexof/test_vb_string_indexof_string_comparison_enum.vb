' vybe-test: vb/vb_string_substring_indexof_lastindexof/test_vb_string_indexof_string_comparison_enum
' origin: languages/vb/tests/vb/test_vb_string_substring_indexof_lastindexof.rs

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
        Dim s As String = "Quick Brown Fox"
        __Check(CStr(s.IndexOf("brown", StringComparison.OrdinalIgnoreCase)), "6")
        __Check(CStr(s.IndexOf("BROWN", StringComparison.OrdinalIgnoreCase)), "6")
    End Sub
End Module
