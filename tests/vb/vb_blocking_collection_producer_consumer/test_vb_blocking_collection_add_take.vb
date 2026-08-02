' vybe-test: vb/vb_blocking_collection_producer_consumer/test_vb_blocking_collection_add_take
' origin: languages/vb/tests/vb/test_vb_blocking_collection_producer_consumer.rs

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
        Dim bc As New BlockingCollection(Of String)()
        bc.Add("Item1")
        bc.Add("Item2")
        __Check(CStr(bc.Take()), "Item1")
        __Check(CStr(bc.Take()), "Item2")
    End Sub
End Module
