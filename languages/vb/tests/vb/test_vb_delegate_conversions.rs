use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Delegate Conversions (Relaxed Delegate Conversion)
// ═══════════════════════════════════════════════════════════

#[test]
fn delegate_conversions_advanced() {
    let out = run_vb(
        r#"
Delegate Sub StringAction(s As String)

Module M
    Sub ExecuteAction(action As StringAction)
        action("Test Message")
    End Sub

    Sub Main()
        ' Relaxed delegate conversion allows ignoring parameters
        Dim action1 As StringAction = AddressOf HandleWithNoArgs
        action1("Ignore me")
        
        ' Or we can pass an anonymous method that takes no arguments
        ExecuteAction(Sub() Console.WriteLine("Action executed with no args"))
    End Sub
    
    Sub HandleWithNoArgs()
        Console.WriteLine("HandleWithNoArgs called")
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec!["HandleWithNoArgs called", "Action executed with no args"]
    );
}
