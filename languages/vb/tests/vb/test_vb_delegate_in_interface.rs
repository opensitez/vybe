use super::helpers::run_vb;

#[test]
fn delegate_in_interface() {
    let out = run_vb(
        r#"
' VB.NET allows defining delegates inside namespaces, modules, or classes, but in Interfaces?
' Interfaces cannot contain nested types in C#, but in VB.NET? Actually VB doesn't allow nested types in interfaces.
' Let's define a delegate outside and an event of that delegate inside.
Delegate Sub MyCustomHandler(msg As String)

Interface INotifier
    Event Notified As MyCustomHandler
End Interface

Class Notifier
    Implements INotifier
    Public Event Notified As MyCustomHandler Implements INotifier.Notified
    
    Public Sub Trigger()
        RaiseEvent Notified("Alert")
    End Sub
End Class

Module M
    Sub Handle(msg As String)
        Console.WriteLine(msg)
    End Sub

    Sub Main()
        Dim n As New Notifier()
        AddHandler n.Notified, AddressOf Handle
        n.Trigger()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Alert"]);
}
