' vybe-test: vb/vb_blocking_collection_take/test_vb_blocking_collection_take_from_any_multiple_collections
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
        Dim bc1 As New BlockingCollection(Of Integer)()
        Dim bc2 As New BlockingCollection(Of Integer)()
        bc2.Add(500)

        Dim item As Integer
        Dim idx = BlockingCollection(Of Integer).TakeFromAny(New BlockingCollection(Of Integer)() {bc1, bc2}, item)
        __Check(CStr("Index: " & idx & "|Item: " & item), "Index: 1|Item: 500")
    End Sub
End Module
