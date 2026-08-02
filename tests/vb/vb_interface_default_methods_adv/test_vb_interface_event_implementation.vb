' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_event_implementation
' origin: languages/vb/tests/vb/test_vb_interface_default_methods_adv.rs

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

Interface INotifier
    Event Notified(msg As String)
    Sub Trigger(msg As String)
End Interface

Class Notifier
    Implements INotifier
    Public Event Notified(msg As String) Implements INotifier.Notified
    Public Sub Trigger(msg As String) Implements INotifier.Trigger
        RaiseEvent Notified(msg)
    End Sub
End Class

Module Program
    Sub Main()
        Dim n As INotifier = New Notifier()
        AddHandler n.Notified, Sub(m) __Check(CStr("RECEIVED: " & m), "RECEIVED: Alert")
        n.Trigger("Alert")
    End Sub
End Module
