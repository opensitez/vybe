use super::helpers::run_vb;

#[test]
fn date_now_today() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' We can't easily assert the exact time, but we can verify types and properties
        Dim n As Date = Now
        Dim t As Date = Today
        
        Console.WriteLine(n.Year >= 2020)
        Console.WriteLine(t.TimeOfDay.TotalSeconds = 0) ' Today has no time component (midnight)
        
        Dim tod As Date = TimeOfDay
        Console.WriteLine(tod.Year = 1) ' TimeOfDay has dummy date component
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}
