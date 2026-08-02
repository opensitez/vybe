' vybe-test: vb/vb_blocking_collection_take/test_vb_blocking_collection_add_to_any_bounded
' origin: languages/vb/tests/vb/test_vb_blocking_collection_take.rs

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
        Dim bc1 As New BlockingCollection(Of Integer)(boundedCapacity:=1)
        Dim bc2 As New BlockingCollection(Of Integer)(boundedCapacity:=1)
        bc1.Add(10)

        ' Adding to collection array routes to bc2 since bc1 is full!
        Dim idx = BlockingCollection(Of Integer).AddToAny(New BlockingCollection(Of Integer)() {bc1, bc2}, 20)
        __Check(CStr("Added to Collection Index: " & idx), "Added to Collection Index: 1")
    End Sub
End Module
