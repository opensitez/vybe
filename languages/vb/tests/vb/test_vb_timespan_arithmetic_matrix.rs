use super::helpers::run_vb;

#[test]
fn timespan_parse_and_fields() {
    let out = run_vb(
        r#"
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
fn timespan_arithmetic_add_subtract() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim t1 As TimeSpan = TimeSpan.FromHours(1)
        Dim t2 As TimeSpan = TimeSpan.FromMinutes(30)
        Dim sum As TimeSpan = t1 + t2
        Dim diff As TimeSpan = t1 - t2
        Console.WriteLine(sum.TotalMinutes)
        Console.WriteLine(diff.TotalMinutes)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["90", "30"]);
}

#[test]
fn timespan_compare_orders_values() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim shortSpan As TimeSpan = TimeSpan.FromMinutes(10)
        Dim longSpan As TimeSpan = TimeSpan.FromMinutes(20)
        Console.WriteLine(shortSpan < longSpan)
        Console.WriteLine(longSpan > shortSpan)
        Console.WriteLine(shortSpan.CompareTo(longSpan) < 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn timespan_from_days_hours_minutes_roundtrip() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim ts As TimeSpan = TimeSpan.FromDays(1)
        ts = ts + TimeSpan.FromHours(2)
        ts = ts + TimeSpan.FromMinutes(30)
        Console.WriteLine(ts.Days)
        Console.WriteLine(ts.Hours)
        Console.WriteLine(ts.Minutes)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "2", "30"]);
}

#[test]
fn timespan_ticks_roundtrip_to_time_span() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim ticks As Long = TimeSpan.FromSeconds(90).Ticks
        Dim ts As New TimeSpan(ticks)
        Console.WriteLine(ts.TotalSeconds)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["90"]);
}

#[test]
fn timespan_negate_flips_sign() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim ts As TimeSpan = TimeSpan.FromMinutes(5)
        Dim neg As TimeSpan = ts.Negate()
        Console.WriteLine(ts.TotalMinutes)
        Console.WriteLine(neg.TotalMinutes)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5", "-5"]);
}

#[test]
fn timespan_duration_is_non_negative() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim negative As TimeSpan = TimeSpan.FromMinutes(-5)
        Console.WriteLine(negative.Duration().TotalMinutes)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5"]);
}

#[test]
fn timespan_zero_is_zero() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(TimeSpan.Zero.Ticks)
        Console.WriteLine(TimeSpan.Zero.TotalSeconds)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0", "0"]);
}
