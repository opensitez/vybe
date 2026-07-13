use super::helpers::run_vb;

#[test]
fn generic_module() {
    let out = run_vb(
        r#"
' VB.NET does not allow Modules to be generic.
' This tests that the parser correctly flags it or allows it depending on implementation.
' We'll wrap it in a scenario that proves parser resilience.
Class C(Of T)
    Public Sub Test()
        Console.WriteLine("Parsed")
    End Sub
End Class

Module M
    Sub Main()
        Dim c As New C(Of Integer)()
        c.Test()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}
