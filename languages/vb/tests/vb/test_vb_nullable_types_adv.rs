use super::helpers::run_vb;

#[test]
fn nullable_types_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Nullable types using ? suffix
        Dim x As Integer? = 10
        Dim y As Integer? = Nothing
        
        Console.WriteLine(x.HasValue)
        Console.WriteLine(y.HasValue)
        
        If x.HasValue Then
            Console.WriteLine(x.Value)
        End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "False", "10"]);
}
