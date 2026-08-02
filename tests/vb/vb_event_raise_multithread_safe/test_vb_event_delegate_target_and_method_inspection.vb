' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_delegate_target_and_method_inspection
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

Class TargetReceiver
    Public Sub OnEvent()
        __Check(CStr("Receiver Action"), "Receiver Action")
    End Sub
End Class

Class EventSource
    Public Event Action As Action
    Public Sub Fire()
        RaiseEvent Action()
    End Sub
End Class

Module Program
    Sub Main()
        Dim tr As New TargetReceiver()
        Dim es As New EventSource()
        AddHandler es.Action, AddressOf tr.OnEvent
        es.Fire()
    End Sub
End Module
