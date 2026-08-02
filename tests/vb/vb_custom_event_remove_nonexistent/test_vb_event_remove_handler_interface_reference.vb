' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_remove_handler_interface_reference
' origin: languages/vb/tests/vb/test_vb_custom_event_remove_nonexistent.rs

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

Interface IEventContainer
    Event Alert As Action
End Interface

Class ConcreteContainer
    Implements IEventContainer
    Public Event Alert As Action Implements IEventContainer.Alert
    Public Sub Fire()
        RaiseEvent Alert()
    End Sub
End Class

Module Program
    Private Sub OnAlert() : __Check(CStr("Alerted"), "Alerted") : End Sub

    Sub Main()
        Dim c As New ConcreteContainer()
        Dim ic As IEventContainer = c
        AddHandler ic.Alert, AddressOf OnAlert
        c.Fire()
        RemoveHandler ic.Alert, AddressOf OnAlert
        c.Fire()
    End Sub
End Module
