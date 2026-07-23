use super::helpers::run_vb;

#[test]
fn datetime_formatting_standard() {
    let out = run_vb(
        r#"
Imports System.Globalization

Module M
    Sub Main()
        Dim d As Date = New Date(2024, 1, 1, 15, 30, 0)
        
        ' Ensure invariant culture for consistent results across environments
        Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture
        
        Console.WriteLine(d.ToString("yyyy-MM-dd"))
        Console.WriteLine(d.ToString("HH:mm:ss"))
        Console.WriteLine(d.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2024-01-01", "15:30:00", "2024-01-01 15:30:00"]);
}

#[test]
fn datetime_parsing() {
    let out = run_vb(
        r#"
Imports System.Globalization

Module M
    Sub Main()
        Dim d1 As Date = Date.Parse("2024-01-01", CultureInfo.InvariantCulture)
        Console.WriteLine(d1.Year)
        
        Dim d2 As Date
        If Date.TryParseExact("20240101", "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, d2) Then
            Console.WriteLine(d2.Month)
        End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2024", "1"]);
}
