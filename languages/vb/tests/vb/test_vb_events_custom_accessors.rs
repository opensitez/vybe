use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Custom Events (Event Accessors)
// ═══════════════════════════════════════════════════════════

#[test]
fn custom_event_accessors() {
    let out = run_vb(
        r#"
Class CustomEventSource
    ' Backing delegate
    Private ActionDelegate As Action
    
    ' Custom Event
    Public Custom Event ActionOccurred As Action
        AddHandler(value As Action)
            ActionDelegate = CType([Delegate].Combine(ActionDelegate, value), Action)
            Console.WriteLine("Handler Added")
        End AddHandler
        
        RemoveHandler(value As Action)
            ActionDelegate = CType([Delegate].Remove(ActionDelegate, value), Action)
            Console.WriteLine("Handler Removed")
        End RemoveHandler
        
        RaiseEvent()
            Console.WriteLine("Raising Event")
            If ActionDelegate IsNot Nothing Then
                ActionDelegate.Invoke()
            End If
        End RaiseEvent
    End Event
    
    Public Sub DoAction()
        RaiseEvent ActionOccurred()
    End Sub
End Class

Module M
    Sub OnAction()
        Console.WriteLine("Action executed")
    End Sub

    Sub Main()
        Dim source As New CustomEventSource()
        AddHandler source.ActionOccurred, AddressOf OnAction
        source.DoAction()
        RemoveHandler source.ActionOccurred, AddressOf OnAction
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Handler Added", "Raising Event", "Action executed", "Handler Removed"]);
}
