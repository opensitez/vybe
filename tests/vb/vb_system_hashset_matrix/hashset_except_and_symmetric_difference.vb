' vybe-test: vb/vb_system_hashset_matrix/hashset_except_and_symmetric_difference
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
        Dim left As New HashSet(Of Integer) From {1, 2, 3, 4}
        left.ExceptWith(New HashSet(Of Integer)({3, 4}))
        __Check(CStr(left.Count), "2")
        Dim a As New HashSet(Of Integer) From {1, 2, 5}
        Dim b As New HashSet(Of Integer) From {2, 3}
        a.SymmetricExceptWith(b)
        __Check(CStr(a.Count), "2")
    End Sub
End Module
