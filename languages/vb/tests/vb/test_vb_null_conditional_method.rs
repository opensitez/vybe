use super::helpers::run_vb;

#[test]
fn null_conditional_method() {
    let out = run_vb(
        r#"
Class Person
    Public Sub DoWork()
        Console.WriteLine("Working")
    End Sub
End Class

Module M
    Sub Main()
        Dim p As Person = Nothing
        
        ' Null conditional method call
        p?.DoWork()
        
        p = New Person()
        p?.DoWork()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Working"]);
}
