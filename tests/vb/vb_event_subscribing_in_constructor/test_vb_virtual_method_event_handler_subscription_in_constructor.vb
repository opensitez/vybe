' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_virtual_method_event_handler_subscription_in_constructor
' origin: languages/vb/tests/vb/test_vb_event_subscribing_in_constructor.rs

Imports System

Class BaseSubscriber
    Public Sub New(pub As Source)
        AddHandler pub.Ping, AddressOf OnPing
    End Sub

    Protected Overridable Sub OnPing(sender As Object, e As EventArgs)
        Console.WriteLine("Base OnPing")
    End Sub
End Class

Class DerivedSubscriber
    Inherits BaseSubscriber

    Public Sub New(pub As Source)
        MyBase.New(pub)
    End Sub

    Protected Overrides Sub OnPing(sender As Object, e As EventArgs)
        Console.WriteLine("Derived OnPing")
    End Sub
End Class

Class Source
    Public Event Ping As EventHandler
    Public Sub Fire()
        RaiseEvent Ping(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As New Source()
        Dim subObj As New DerivedSubscriber(s)
        s.Fire()
    End Sub
End Module
