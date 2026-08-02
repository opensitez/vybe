' vybe-test: vb/vb_spec_delegates_lambdas/delegate_spec_removehandler_with_lambda_variable_detaches_listener
' origin: languages/vb/tests/vb/test_vb_spec_delegates_lambdas.rs

Class Clock
    Public Event Tick()
    Public Sub RaiseTick()
        RaiseEvent Tick()
    End Sub
End Class
Module M
    Sub Main()
        Dim clock As New Clock()
        Dim handler As Action = Sub() Console.WriteLine("tick")
        AddHandler clock.Tick, handler
        RemoveHandler clock.Tick, handler
        clock.RaiseTick()
        Console.WriteLine("done")
    End Sub
End Module
