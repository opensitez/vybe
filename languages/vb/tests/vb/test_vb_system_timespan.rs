use super::helpers::run_vb;

#[test]
fn system_timespan_parsing() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim ts As TimeSpan = TimeSpan.Parse("01:30:00")
        Console.WriteLine(ts.Hours)
        Console.WriteLine(ts.Minutes)
        Console.WriteLine(ts.TotalMinutes)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1", "30", "90"]);
}

#[test]
fn system_timespan_arithmetic() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim ts1 As New TimeSpan(1, 0, 0)
        Dim ts2 As New TimeSpan(0, 30, 0)
        
        Dim ts3 = ts1 + ts2
        Console.WriteLine(ts3.TotalMinutes)
        
        Dim ts4 = ts1 - ts2
        Console.WriteLine(ts4.TotalMinutes)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["90", "30"]);
}
