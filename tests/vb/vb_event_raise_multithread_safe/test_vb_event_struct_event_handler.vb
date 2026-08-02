' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_struct_event_handler
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

Structure ValuePublisher
    Public Event ValueChanged As Action(Of Integer)
    Public Sub Publish(val As Integer)
        RaiseEvent ValueChanged(val)
    End Sub
End Structure

Module Program
    Sub Main()
        Dim vp As New ValuePublisher()
        AddHandler vp.ValueChanged, Sub(v) __Check(CStr("Val: " & v), "Val: 99")
        vp.Publish(99)
    End Sub
End Module
