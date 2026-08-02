' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_withevents_single_method_handles_multiple_events
' origin: languages/vb/tests/vb/test_vb_event_subscribing_in_constructor.rs

Imports System

Class Source
    Public Event Event1 As EventHandler
    Public Event Event2 As EventHandler
    Public Sub Fire1()
        RaiseEvent Event1(Me, EventArgs.Empty)
    End Sub
    Public Sub Fire2()
        RaiseEvent Event2(Me, EventArgs.Empty)
    End Sub
End Class

Class MultiHandleListener
    Public WithEvents Src As Source

    ' Single handler method handles both Event1 and Event2!
    Private Sub OnCombined(sender As Object, e As EventArgs) Handles Src.Event1, Src.Event2
        Console.WriteLine("Combined Event Handled")
    End Sub
End Class

Module Program
    Sub Main()
        Dim l As New MultiHandleListener With {.Src = New Source()}
        l.Src.Fire1()
        l.Src.Fire2()
    End Sub
End Module
