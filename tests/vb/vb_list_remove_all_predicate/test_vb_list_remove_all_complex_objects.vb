' vybe-test: vb/vb_list_remove_all_predicate/test_vb_list_remove_all_complex_objects
' origin: languages/vb/tests/vb/test_vb_list_remove_all_predicate.rs

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

Class TaskItem
    Public Property Title As String
    Public Property IsDone As Boolean
    Public Sub New(t As String, done As Boolean)
        Title = t : IsDone = done
    End Sub
End Class

Module Program
    Sub Main()
        Dim tasks As New List(Of TaskItem) From {
            New TaskItem("T1", True),
            New TaskItem("T2", False),
            New TaskItem("T3", True)
        }
        tasks.RemoveAll(Function(t) t.IsDone)
        __Check(CStr(tasks.Count & ":" & tasks(0).Title), "1:T2")
    End Sub
End Module
