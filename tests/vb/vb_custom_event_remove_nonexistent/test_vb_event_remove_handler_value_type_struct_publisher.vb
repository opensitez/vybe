' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_remove_handler_value_type_struct_publisher
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

Structure StructPublisher
    Public Event Signal As Action
    Public Sub Fire()
        RaiseEvent Signal()
    End Sub
End Structure

Module Program
    Private Sub OnSignal() : __Check(CStr("Signal"), "Signal") : End Sub

    Sub Main()
        Dim sp As New StructPublisher()
        AddHandler sp.Signal, AddressOf OnSignal
        sp.Fire()
        RemoveHandler sp.Signal, AddressOf OnSignal
        sp.Fire()
    End Sub
End Module
