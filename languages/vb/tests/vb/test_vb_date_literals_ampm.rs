use super::helpers::run_vb;

#[test]
fn date_literals_ampm() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Date literals support AM/PM 
        Dim d1 As Date = #8:30:00 AM#
        Dim d2 As Date = #1/1/2000 8:30:00 PM#
        
        Console.WriteLine(d1.Hour)
        Console.WriteLine(d2.Hour)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["8", "20"]);
}
