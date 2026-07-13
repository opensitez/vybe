use super::helpers::run_vb;

#[test]
fn operator_precedence_math() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' ^ has higher precedence than *, /
        Console.WriteLine(2 + 3 * 4 ^ 2) ' 2 + 3 * 16 = 2 + 48 = 50
        
        ' \ (integer division) has lower precedence than *, / but higher than +, -
        Console.WriteLine(10 + 20 \ 3) ' 10 + 6 = 16
        
        ' Mod has lower precedence than \
        Console.WriteLine(10 Mod 3 + 1) ' 1 + 1 = 2
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["50", "16", "2"]);
}

#[test]
fn operator_precedence_logical() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Not has highest, then And, then Or
        Console.WriteLine(True Or False And Not True) ' True Or (False And False) = True
        Console.WriteLine((True Or False) And Not True) ' True And False = False
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}
