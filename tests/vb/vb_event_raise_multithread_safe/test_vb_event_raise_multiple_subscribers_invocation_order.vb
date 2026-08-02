' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_raise_multiple_subscribers_invocation_order
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

Class Emitter
    Public Event Trigger As Action(Of Integer)
    Public Sub Fire(val As Integer)
        RaiseEvent Trigger(val)
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New Emitter()
        AddHandler e.Trigger, Sub(v) __Check(CStr("Sub1: " & v), "Sub1: 10")
        AddHandler e.Trigger, Sub(v) __Check(CStr("Sub2: " & (v * 2)), "Sub2: 20")
        e.Fire(10)
    End Sub
End Module
