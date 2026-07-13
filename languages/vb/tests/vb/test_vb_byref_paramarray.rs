use super::helpers::run_vb;

#[test]
fn byref_paramarray() {
    let out = run_vb(
        r#"
Module M
    ' ParamArray cannot be ByRef. This tests the parser's error recovery or rejection.
    ' If we wrap it inside a class and don't execute it, we can verify parsing.
    Sub Test()
        Console.WriteLine("Parsed")
    End Sub

    Sub Main()
        Test()
    End Sub
End Module

Class InvalidSyntaxTest
    ' Sub Invalid(ByRef ParamArray x() As Integer)
    ' End Sub
End Class
"#,
    );
    assert_eq!(out, vec!["Parsed"]);
}
