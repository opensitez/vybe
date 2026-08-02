' vybe-test: vb/vb_blocking_collection_take/test_vb_blocking_collection_bounded_capacity
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
        Dim bc As New BlockingCollection(Of Integer)(boundedCapacity:=2)
        Dim added1 = bc.TryAdd(1)
        Dim added2 = bc.TryAdd(2)
        Dim added3 = bc.TryAdd(3, millisecondsTimeout:=10) ' Exceeds bounded capacity!

        __Check(CStr(added1 & "|" & added2 & "|" & added3), "True|True|False")
    End Sub
End Module
