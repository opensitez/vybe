' vybe-test: vb/vb_concurrent_bag_operations/test_vb_concurrent_bag_to_array_snapshot
' origin: languages/vb/tests/vb/test_vb_concurrent_bag_operations.rs

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

Module Program
    Sub Main()
        Dim bag As New ConcurrentBag(Of Integer)()
        bag.Add(1)
        bag.Add(2)
        Dim arr As Integer() = bag.ToArray()
        __Check(CStr(arr.Length), "2")
    End Sub
End Module
