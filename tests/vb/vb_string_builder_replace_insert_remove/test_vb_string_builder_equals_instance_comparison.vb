' vybe-test: vb/vb_string_builder_replace_insert_remove/test_vb_string_builder_equals_instance_comparison
' origin: languages/vb/tests/vb/test_vb_string_builder_replace_insert_remove.rs

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

Imports System.Text

Module Program
    Sub Main()
        Dim sb1 As New StringBuilder("Text")
        Dim sb2 As New StringBuilder("Text")
        ' StringBuilder.Equals checks structural capacity/content equality!
        __Check(CStr(sb1.Equals(sb2)), "True")
    End Sub
End Module
