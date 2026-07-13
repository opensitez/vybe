use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Late Binding (Methods)
// ═══════════════════════════════════════════════════════════

#[test]
fn late_binding_method_call() {
    let out = run_vb(
        r#"
Class Greeter
    Public Function SayHello(name As String) As String
        Return "Hello " & name
    End Function
End Class

Module M
    Sub Main()
        ' Using Object type forces late binding (if Option Strict is Off, which is default)
        Dim g As Object = New Greeter()
        
        ' Method call is resolved at runtime
        Console.WriteLine(g.SayHello("VB"))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello VB"]);
}

#[test]
fn late_binding_with_arguments() {
    let out = run_vb(
        r#"
Class MathOperations
    Public Function Add(a As Integer, b As Integer) As Integer
        Return a + b
    End Function
End Class

Module M
    Sub Main()
        Dim op As Object = New MathOperations()
        Dim result As Object = op.Add(10, 20)
        Console.WriteLine(result)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["30"]);
}
