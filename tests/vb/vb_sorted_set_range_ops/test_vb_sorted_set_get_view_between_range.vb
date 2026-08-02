' vybe-test: vb/vb_sorted_set_range_ops/test_vb_sorted_set_get_view_between_range
' origin: languages/vb/tests/vb/test_vb_sorted_set_range_ops.rs

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

Module Program
    Sub Main()
        Dim ss As New SortedSet(Of Integer) From {10, 20, 30, 40, 50, 60}
        Dim view As SortedSet(Of Integer) = ss.GetViewBetween(20, 50)
        __Check(CStr(String.Join(",", view)), "20,30,40,50")
    End Sub
End Module
