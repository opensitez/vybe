use super::helpers::run_vb;

#[test]
fn custom_event_blocks() {
    let out = run_vb(
        r#"
Class EventPublisher
    ' A Custom Event allows defining AddHandler, RemoveHandler, and RaiseEvent blocks
    Public Custom Event MyEvent As EventHandler
        AddHandler(value As EventHandler)
            Console.WriteLine("Handler Added")
        End AddHandler
        RemoveHandler(value As EventHandler)
            Console.WriteLine("Handler Removed")
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            Console.WriteLine("Event Raised")
        End RaiseEvent
    End Event
    
    Public Sub Trigger()
        RaiseEvent MyEvent(Me, EventArgs.Empty)
    End Sub
End Class

Module M
    Sub Handler(sender As Object, e As EventArgs)
    End Sub

    Sub Main()
        Dim p As New EventPublisher()
        AddHandler p.MyEvent, AddressOf Handler
        p.Trigger()
        RemoveHandler p.MyEvent, AddressOf Handler
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec!["Handler Added", "Event Raised", "Handler Removed"]
    );
}
