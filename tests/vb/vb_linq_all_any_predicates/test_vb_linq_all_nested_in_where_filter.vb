' vybe-test: vb/vb_linq_all_any_predicates/test_vb_linq_all_nested_in_where_filter
' origin: languages/vb/tests/vb/test_vb_linq_all_any_predicates.rs

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
Imports System.Linq

Class Team
    Public Property Scores As List(Of Integer)
End Class

Module Program
    Sub Main()
        Dim teams As New List(Of Team) From {
            New Team With {.Scores = New List(Of Integer) From {10, 20, 30}},
            New Team With {.Scores = New List(Of Integer) From {25, 35, 45}}
        }
        Dim highScorers = teams.Where(Function(t) t.Scores.All(Function(s) s >= 20))
        __Check(CStr(highScorers.Count()), "1")
    End Sub
End Module
