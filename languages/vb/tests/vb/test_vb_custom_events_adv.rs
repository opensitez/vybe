use super::helpers::run_vb;

#[test]
fn custom_events_generic_delegate() {
    let out = run_vb(
        r#"
Class Subject
    Private _handlers As EventHandler(Of String)
    
    Public Custom Event Notify As EventHandler(Of String)
        AddHandler(value As EventHandler(Of String))
            _handlers = CType([Delegate].Combine(_handlers, value), EventHandler(Of String))
        End AddHandler
        RemoveHandler(value As EventHandler(Of String))
            _handlers = CType([Delegate].Remove(_handlers, value), EventHandler(Of String))
        End RemoveHandler
        RaiseEvent(sender As Object, e As String)
            If _handlers IsNot Nothing Then
                _handlers.Invoke(sender, e)
            End If
        End RaiseEvent
    End Event
    
    Public Sub Trigger(msg As String)
        RaiseEvent Notify(Me, msg)
    End Sub
End Class

Module M
    Sub OnNotify(sender As Object, e As String)
        Console.WriteLine("Notified: " & e)
    End Sub

    Sub Main()
        Dim s As New Subject()
        AddHandler s.Notify, AddressOf OnNotify
        s.Trigger("Hello")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Notified: Hello"]);
}
