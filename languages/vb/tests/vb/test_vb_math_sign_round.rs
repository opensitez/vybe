use super::helpers::run_vb;

#[test]
fn math_sign_round() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Sign returns 1, 0, or -1
        Console.WriteLine(Sign(-42))
        Console.WriteLine(Sign(0))
        Console.WriteLine(Sign(42))
        
        ' Round performs banker's rounding by default
        Console.WriteLine(Round(2.5)) ' Rounds to nearest even -> 2
        Console.WriteLine(Round(3.5)) ' Rounds to nearest even -> 4
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["-1", "0", "1", "2", "4"]);
}
