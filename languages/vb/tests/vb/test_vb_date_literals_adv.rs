use super::helpers::run_vb;

#[test]
fn date_literals_formats() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Multiple formats allowed in date literals
        Dim d1 As Date = #1998-11-23#
        Dim d2 As Date = #23 Nov 98#
        Dim d3 As Date = #1:15 PM#
        
        Console.WriteLine(d1.Year)
        Console.WriteLine(d2.Month)
        Console.WriteLine(d3.Hour)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1998", "11", "13"]);
}
