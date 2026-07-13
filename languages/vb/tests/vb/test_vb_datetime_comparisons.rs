use super::helpers::run_vb;

#[test]
fn datetime_comparisons() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim d1 As Date = #2024-01-01 12:00:00#
        Dim d2 As Date = #2024-01-01 12:00:00#
        Dim d3 As Date = #2024-01-02#
        
        Console.WriteLine(d1 = d2)
        Console.WriteLine(d1 < d3)
        Console.WriteLine(d3 >= d1)
        
        ' Compare method
        Console.WriteLine(Date.Compare(d1, d3))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "True", "True", "-1"]);
}

#[test]
fn datetime_parts() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim d As Date = #2024-02-15 14:30:45#
        
        Console.WriteLine(d.Year)
        Console.WriteLine(d.Month)
        Console.WriteLine(d.Day)
        Console.WriteLine(d.Hour)
        Console.WriteLine(d.Minute)
        Console.WriteLine(d.Second)
        Console.WriteLine(d.DayOfWeek.ToString())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2024", "2", "15", "14", "30", "45", "Thursday"]);
}
