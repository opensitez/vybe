' vybe-test: vb/vb_blocking_collection_take/test_vb_blocking_collection_take_from_any_timeout_returns_minus_one
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

        Dim item As Integer
        ' Timeout after 10ms when both collections are empty
        Dim idx = BlockingCollection(Of Integer).TryTakeFromAny(New BlockingCollection(Of Integer)() {bc1, bc2}, item, millisecondsTimeout:=10)
        __Check(CStr(idx & "|" & item), "-1|0")
    End Sub
End Module
