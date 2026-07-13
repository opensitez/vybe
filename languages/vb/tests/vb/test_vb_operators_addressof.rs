use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Operators (AddressOf)
// ═══════════════════════════════════════════════════════════

#[test]
fn operator_addressof() {
    let out = run_vb(
        r#"
Module M
    Sub PrintMessage(msg As String)
        Console.WriteLine("Message: " & msg)
    End Sub

    Sub Main()
        ' AddressOf creates a delegate pointing to the specified procedure
        Dim del As Action(Of String) = AddressOf PrintMessage
        del("Hello World")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Message: Hello World"]);
}
