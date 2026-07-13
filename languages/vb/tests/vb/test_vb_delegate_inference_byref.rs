use super::helpers::run_vb;

#[test]
fn delegate_inference_byref() {
    let out = run_vb(
        r#"
Module M
    Delegate Sub ByRefAction(ByRef x As Integer)

    Sub Main()
        ' Delegate type inference with ByRef
        Dim act As ByRefAction = Sub(ByRef x As Integer) x += 1
        
        Dim val = 10
        act(val)
        Console.WriteLine(val)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["11"]);
}
