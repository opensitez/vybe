use super::helpers::run_vb;

#[test]
fn partial_method_byref() {
    let out = run_vb(
        r#"
Partial Class Processor
    Partial Private Sub ModifyValue(ByRef val As Integer)
    End Sub
    
    Public Sub Run()
        Dim x = 10
        ModifyValue(x)
        Console.WriteLine(x)
    End Sub
End Class

Partial Class Processor
    Private Sub ModifyValue(ByRef val As Integer)
        val = 20
    End Sub
End Class

Module M
    Sub Main()
        Dim p As New Processor()
        p.Run()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["20"]);
}
