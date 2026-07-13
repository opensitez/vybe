use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Delegates (Relaxed Instantiation)
// ═══════════════════════════════════════════════════════════

#[test]
fn delegate_relaxed_instantiation() {
    let out = run_vb(
        r#"
Module M
    Sub PrintMessage(msg As String)
        Console.WriteLine(msg)
    End Sub

    ' Delegate requires an Object and EventArgs
    Delegate Sub EventHandler(sender As Object, e As EventArgs)

    Sub Main()
        ' Relaxed delegate instantiation allows dropping parameters if the target doesn't need them
        ' OR passing arguments that can be widened/narrowed automatically
        ' A common VB feature is assigning a Sub with no parameters to an EventHandler
        Dim handler As EventHandler = AddressOf LogEvent
        handler(Nothing, Nothing)
    End Sub

    Sub LogEvent()
        Console.WriteLine("Event Logged without parameters")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Event Logged without parameters"]);
}

#[test]
fn delegate_return_type_relaxation() {
    let out = run_vb(
        r#"
Module M
    Delegate Function Provider() As Object

    Function ProvideString() As String
        Return "Hello"
    End Function

    Sub Main()
        ' String narrows/widens to Object, so this is valid relaxed binding
        Dim p As Provider = AddressOf ProvideString
        Console.WriteLine(p().ToString())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello"]);
}
