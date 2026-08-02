' vybe-test: vb/vb_system_textbuilder_matrix/textbuilder_replace_targets_all_matches
' origin: languages/vb/tests/vb/test_vb_system_textbuilder_matrix.rs

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

Module M
    Sub Main()
        Dim sb As New StringBuilder("banana")
        sb.Replace("a", "o")
        __Check(CStr(sb.ToString()), "bonono")
    End Sub
End Module
