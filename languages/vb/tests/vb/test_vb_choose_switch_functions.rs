use super::helpers::run_vb;

#[test]
fn choose_function() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Choose returns a value from a list of arguments based on an index (1-based)
        Dim val1 = Choose(1, "Apple", "Banana", "Cherry")
        Dim val2 = Choose(3, "Apple", "Banana", "Cherry")
        
        Console.WriteLine(val1)
        Console.WriteLine(val2)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Apple", "Cherry"]);
}

#[test]
fn switch_function() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim age As Integer = 25
        
        ' Switch evaluates a list of expressions and returns the corresponding value for the first True expression
        Dim category = Switch(
            age < 18, "Minor",
            age >= 18 AndAlso age < 65, "Adult",
            age >= 65, "Senior"
        )
        
        Console.WriteLine(category)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Adult"]);
}
