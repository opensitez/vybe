' vybe-test: vb/vb_system_concurrent_collections_matrix/concurrent_bag_add_take_and_count
' origin: languages/vb/tests/vb/test_vb_system_concurrent_collections_matrix.rs

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

Imports System.Collections.Concurrent

Module M
    Sub Main()
        Dim bag As New ConcurrentBag(Of String)()
        bag.Add("left")
        bag.Add("right")

        Dim countBefore As Integer = bag.Count
        Dim head As String = ""

        __Check(CStr(countBefore), "2")
        __Check(CStr(bag.TryPeek(head)), "True")
        __Check(CStr(head = "left" OrElse head = "right"), "True")

        Dim taken As String = ""
        __Check(CStr(bag.TryTake(taken)), "True")
        __Check(CStr(bag.Count), "1")
        __Check(CStr(taken = "left" OrElse taken = "right"), "True")
    End Sub
End Module
