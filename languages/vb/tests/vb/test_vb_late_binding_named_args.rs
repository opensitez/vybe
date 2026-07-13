use super::helpers::run_vb;

#[test]
fn late_binding_named_args() {
    let out = run_vb(
        r#"
Option Strict Off

Class Target
    Public Sub Process(Arg As Integer)
        Console.WriteLine(Arg)
    End Sub
End Class

Module M
    Sub Main()
        Dim obj As Object = New Target()
        
        ' Late binding with named arguments
        obj.Process(Arg:=42)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42"]);
}
