use super::helpers::run_vb;

#[test]
fn delegates_custom_events() {
    let out = run_vb(
        r#"
Class Timer
    Private _tickHandlers As System.EventHandler
    
    ' Custom event accessors
    Public Custom Event Tick As System.EventHandler
        AddHandler(value As System.EventHandler)
            _tickHandlers = CType([Delegate].Combine(_tickHandlers, value), System.EventHandler)
            Console.WriteLine("Handler added")
        End AddHandler
        RemoveHandler(value As System.EventHandler)
            _tickHandlers = CType([Delegate].Remove(_tickHandlers, value), System.EventHandler)
            Console.WriteLine("Handler removed")
        End RemoveHandler
        RaiseEvent(sender As Object, e As System.EventArgs)
            If _tickHandlers IsNot Nothing Then
                _tickHandlers.Invoke(sender, e)
            End If
        End RaiseEvent
    End Event
    
    Public Sub DoTick()
        RaiseEvent Tick(Me, System.EventArgs.Empty)
    End Sub
End Class

Module M
    Sub OnTick(sender As Object, e As System.EventArgs)
        Console.WriteLine("Tick occurred")
    End Sub

    Sub Main()
        Dim t As New Timer()
        AddHandler t.Tick, AddressOf OnTick
        t.DoTick()
        RemoveHandler t.Tick, AddressOf OnTick
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec!["Handler added", "Tick occurred", "Handler removed"]
    );
}
