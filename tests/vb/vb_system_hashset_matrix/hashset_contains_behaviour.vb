' vybe-test: vb/vb_system_hashset_matrix/hashset_contains_behaviour
' origin: languages/vb/tests/vb/test_vb_system_hashset_matrix.rs

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

Imports System.Collections.Generic

Module M
    Sub Main()
        Dim set As New HashSet(Of String)()
        set.Add("a")
        set.Add("b")
        __Check(CStr(set.Contains("a")), "True")
        __Check(CStr(set.Contains("c")), "False")
    End Sub
End Module
