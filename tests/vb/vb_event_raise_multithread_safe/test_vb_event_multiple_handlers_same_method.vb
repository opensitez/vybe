' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_multiple_handlers_same_method
' origin: languages/vb/tests/vb/test_vb_event_raise_multithread_safe.rs

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

Imports System

Class Counter
    Public Event Increment As Action
    Public Sub Count()
        RaiseEvent Increment()
    End Sub
End Class

Module Program
    Private total As Integer = 0
    Private Sub AddOne()
        total += 1
    End Sub

    Sub Main()
        Dim c As New Counter()
        AddHandler c.Increment, AddressOf AddOne
        AddHandler c.Increment, AddressOf AddOne
        c.Count()
        __Check(CStr(total), "2")
    End Sub
End Module
