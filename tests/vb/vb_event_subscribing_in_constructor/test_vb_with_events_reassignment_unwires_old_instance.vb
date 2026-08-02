' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_with_events_reassignment_unwires_old_instance
' origin: languages/vb/tests/vb/test_vb_event_subscribing_in_constructor.rs

Imports System

Class Emitter
    Public Property Name As String
    Public Event Action As EventHandler
    Public Sub Fire()
        RaiseEvent Action(Me, EventArgs.Empty)
    End Sub
End Class

Class SwitchableListener
    Private WithEvents currentEmitter As Emitter

    Public Sub SetEmitter(e As Emitter)
        currentEmitter = e ' Unwires previous currentEmitter, wires new e!
    End Sub

    Private Sub OnAction(sender As Object, e As EventArgs) Handles currentEmitter.Action
        Console.WriteLine("Action Handled From: " & currentEmitter.Name)
    End Sub
End Class

Module Program
    Sub Main()
        Dim e1 As New Emitter With {.Name = "First"}
        Dim e2 As New Emitter With {.Name = "Second"}

        Dim listener As New SwitchableListener()
        listener.SetEmitter(e1)
        e1.Fire()

        listener.SetEmitter(e2)
        e1.Fire() ' Should NOT fire listener!
        e2.Fire() ' Should fire listener!
    End Sub
End Module
