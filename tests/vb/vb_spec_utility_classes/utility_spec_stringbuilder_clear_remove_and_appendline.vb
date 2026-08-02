' vybe-test: vb/vb_spec_utility_classes/utility_spec_stringbuilder_clear_remove_and_appendline
' origin: languages/vb/tests/vb/test_vb_spec_utility_classes.rs

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
        Dim sb As New StringBuilder()
        sb.AppendLine("alpha")
        sb.Append("beta")
        sb.Remove(0, 6)
        __Check(CStr(sb.ToString()), "beta")
        sb.Clear()
        __Check(CStr(sb.Length), "0")
    End Sub
End Module
