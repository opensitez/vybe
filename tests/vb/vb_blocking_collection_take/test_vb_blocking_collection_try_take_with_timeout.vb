' vybe-test: vb/vb_blocking_collection_take/test_vb_blocking_collection_try_take_with_timeout
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
        Dim bc As New BlockingCollection(Of Integer)()
        Dim item As Integer
        Dim ok = bc.TryTake(item, millisecondsTimeout:=10)
        __Check(CStr(ok & "|" & item), "False|0")
    End Sub
End Module
