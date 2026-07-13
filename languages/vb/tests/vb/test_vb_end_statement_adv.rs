use super::helpers::run_vb;

#[test]
fn end_statement_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine("Before End")
        ' End statement terminates execution immediately
        ' Some test runners might fail if we actually execute End, 
        ' so we place it inside an unreachable block
        If False Then
            End
        End If
        Console.WriteLine("Parsed End")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Before End", "Parsed End"]);
}
