use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: AddressOf with Lambda
// ═══════════════════════════════════════════════════════════

#[test]
fn addressof_lambda() {
    let out = run_vb(
        r#"
Module M
    Sub Execute(action As Action)
        action()
    End Sub

    Sub Main()
        ' While lambdas don't strictly need AddressOf,
        ' sometimes they are used with delegates
        Dim a As Action = Sub() Console.WriteLine("Action executed")
        Execute(a)
        
        ' AddressOf can sometimes be used to refer to named methods and passed where a delegate is expected
        Execute(AddressOf PrintMessage)
    End Sub
    
    Sub PrintMessage()
        Console.WriteLine("Message printed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Action executed", "Message printed"]);
}
