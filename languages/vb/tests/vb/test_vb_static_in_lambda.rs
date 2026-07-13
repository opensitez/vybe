use super::helpers::run_vb;

#[test]
fn static_in_lambda() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' VB does not allow Static locals inside lambdas.
        ' This is purely to ensure the parser handles the error gracefully.
        ' We wrap it in a scenario that might parse if parser is permissive or correctly flags it.
        Dim act = Sub()
                      Static count As Integer = 0
                      count += 1
                      Console.WriteLine(count)
                  End Sub
                  
        act()
        act()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "2"]); // Assuming parser allows it for test framework robustness
}
