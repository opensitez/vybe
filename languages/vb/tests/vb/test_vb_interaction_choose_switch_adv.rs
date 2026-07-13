use super::helpers::run_vb;

#[test]
fn interaction_choose_switch_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim index As Integer = 2
        
        ' Choose evaluates index and returns the 1-based matching argument
        Dim choice = Choose(index, "A", "B", "C")
        Console.WriteLine(choice)
        
        ' Switch returns the value associated with the first True expression
        Dim val As Integer = 10
        Dim category = Switch(
            val < 0, "Negative",
            val = 0, "Zero",
            val > 0, "Positive"
        )
        Console.WriteLine(category)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["B", "Positive"]);
}
