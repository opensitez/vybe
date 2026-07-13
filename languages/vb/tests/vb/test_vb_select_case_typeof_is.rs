use super::helpers::run_vb;

#[test]
fn select_case_typeof_is() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim obj As Object = "Hello"
        
        ' Note: Select Case TypeOf is NOT valid syntax in VB.NET.
        ' Instead, you use If/ElseIf TypeOf obj Is Type or pattern matching.
        ' But let's verify parser behaviour for a standard Select Case with boolean expressions.
        Select Case True
            Case TypeOf obj Is String
                Console.WriteLine("String")
            Case TypeOf obj Is Integer
                Console.WriteLine("Integer")
        End Select
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["String"]);
}
