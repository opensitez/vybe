use super::helpers::run_vb;

#[test]
fn date_add_diff() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim d1 As Date = #1/1/2020#
        
        ' DateAdd
        Dim d2 = DateAdd(DateInterval.Day, 10, d1)
        Console.WriteLine(d2.Day)
        
        ' DateDiff
        Dim diff = DateDiff(DateInterval.Day, d1, d2)
        Console.WriteLine(diff)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["11", "10"]);
}
