use super::helpers::run_vb;

#[test]
fn isnothing_operator() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim obj As Object = Nothing
        Dim obj2 As New Object()
        
        ' Legacy IsNothing function vs Is Nothing operator
        Console.WriteLine(IsNothing(obj))
        Console.WriteLine(obj Is Nothing)
        
        Console.WriteLine(IsNothing(obj2))
        Console.WriteLine(obj2 IsNot Nothing)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "True", "False", "True"]);
}
