use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Custom Event Blocks (AddHandler, RemoveHandler, RaiseEvent)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_custom_event_block_explicit_accessors() {
    let src = r#"
Imports System

Public Delegate Sub CustomHandler(msg As String)

Class Publisher
    Private _handlers As CustomHandler

    Public Custom Event StateChanged As CustomHandler
        AddHandler(value As CustomHandler)
            Console.WriteLine("Added")
            _handlers = CType([Delegate].Combine(_handlers, value), CustomHandler)
        End AddHandler

        RemoveHandler(value As CustomHandler)
            Console.WriteLine("Removed")
            _handlers = CType([Delegate].Remove(_handlers, value), CustomHandler)
        End RemoveHandler

        RaiseEvent(msg As String)
            Console.WriteLine("Raising")
            _handlers?.Invoke(msg)
        End RaiseEvent
    End Event

    Public Sub Trigger(msg As String)
        RaiseEvent StateChanged(msg)
    End Sub
End Class

Module Program
    Sub Main()
        Dim pub As New Publisher()
        Dim h As CustomHandler = Sub(m) Console.WriteLine("Received: " & m)
        AddHandler pub.StateChanged, h
        pub.Trigger("Data1")
        RemoveHandler pub.StateChanged, h
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Added", "Raising", "Received: Data1", "Removed"]
    );
}
