use super::helpers::run_vb;

#[test]
fn datetime_from_parts_and_ticks() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim dt As New DateTime(2026, 7, 21, 13, 40, 0)
        Console.WriteLine(dt.Year)
        Console.WriteLine(dt.Month)
        Console.WriteLine(dt.Day)
        Console.WriteLine(dt.Hour)
        Console.WriteLine(dt.Minute)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2026", "7", "21", "13", "40"]);
}

#[test]
fn datetime_adds_and_comparisons_are_expected() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim base As DateTime = New DateTime(2026, 1, 1)
        Dim plusDays As DateTime = base.AddDays(1)
        Dim plusMonths As DateTime = base.AddMonths(1)

        Console.WriteLine(base < plusDays)
        Console.WriteLine(plusDays.Day)
        Console.WriteLine(plusMonths.Month)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "2", "2"]);
}

#[test]
fn datetime_parse_and_try_parse_exact() {
    let out = run_vb(
        r#"
Imports System.Globalization

Module M
    Sub Main()
        Dim parsed As DateTime = DateTime.Parse("2026-07-21")
        Dim ok As Boolean
        Dim exact As DateTime

        ok = DateTime.TryParseExact("21/07/2026", "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, exact)

        Console.WriteLine(parsed.Year)
        Console.WriteLine(ok)
        Console.WriteLine(exact.Month)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2026", "True", "7"]);
}

#[test]
fn datetime_kind_and_utc_roundtrip() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim utc As DateTime = DateTime.SpecifyKind(New DateTime(2026, 7, 21, 0, 0, 0), DateTimeKind.Utc)
        Dim local As DateTime = utc.ToLocalTime()

        Console.WriteLine(utc.Kind = DateTimeKind.Utc)
        Console.WriteLine(local.Kind <> DateTimeKind.Unspecified)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn datetime_ticks_roundtrip() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim source As DateTime = DateTime.Now
        Dim ticks As Long = source.Ticks
        Dim roundTrip As New DateTime(ticks)

        Console.WriteLine(source.Ticks = roundTrip.Ticks)
        Console.WriteLine(source.Year = roundTrip.Year)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}
