use super::helpers::run_vb;

#[test]
fn interaction_iif_vs_if() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim condition As Boolean = True
        
        ' IIf is a legacy function that evaluates BOTH true and false arguments
        ' (Not short-circuited!)
        Dim result1 = IIf(condition, "Yes", "No")
        Console.WriteLine(result1)
        
        ' If operator is short-circuited and type-safe
        Dim result2 = If(condition, "Yes", "No")
        Console.WriteLine(result2)
        
        ' If operator with two arguments acts like coalesce (expr1 ?? expr2)
        Dim val1 As String = Nothing
        Dim val2 As String = "Default"
        Console.WriteLine(If(val1, val2))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Yes", "Yes", "Default"]);
}
