' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_withevents_subclass_overrides_handler_method
' origin: languages/vb/tests/vb/test_vb_event_subscribing_in_constructor.rs

Imports System

Class Publisher
    Public Event Trigger As EventHandler
    Public Sub Fire()
        RaiseEvent Trigger(Me, EventArgs.Empty)
    End Sub
End Class

Class BaseListener
    Public WithEvents Pub As Publisher

    Protected Overridable Sub OnTrigger(sender As Object, e As EventArgs) Handles Pub.Trigger
        Console.WriteLine("Base OnTrigger")
    End Sub
End Class

Class OverridingListener
    Inherits BaseListener

    Protected Overrides Sub OnTrigger(sender As Object, e As EventArgs)
        Console.WriteLine("Overridden OnTrigger")
    End Sub
End Class

Module Program
    Sub Main()
        Dim ol As New OverridingListener()
        Dim p As New Publisher()
        ol.Pub = p
        p.Fire()
    End Sub
End Module
