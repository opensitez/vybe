use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: TimeSpan Properties, Conversions & Arithmetic
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_time_span_from_days() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ts = TimeSpan.FromDays(2.5)
        Console.WriteLine(ts.Days & "|" & ts.Hours & "|" & ts.TotalHours)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|12|60"]);
}

#[test]
fn test_vb_time_span_from_hours() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ts = TimeSpan.FromHours(1.5)
        Console.WriteLine(ts.Hours & ":" & ts.Minutes & "|" & ts.TotalMinutes)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:30|90"]);
}

#[test]
fn test_vb_time_span_from_minutes() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ts = TimeSpan.FromMinutes(90)
        Console.WriteLine(ts.Hours & ":" & ts.Minutes)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:30"]);
}

#[test]
fn test_vb_time_span_from_seconds() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ts = TimeSpan.FromSeconds(3661)
        Console.WriteLine(ts.Hours & ":" & ts.Minutes & ":" & ts.Seconds)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:1:1"]);
}

#[test]
fn test_vb_time_span_from_milliseconds() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ts = TimeSpan.FromMilliseconds(1500)
        Console.WriteLine(ts.Seconds & "." & ts.Milliseconds)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1.500"]);
}

#[test]
fn test_vb_time_span_from_ticks() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ts = TimeSpan.FromTicks(TimeSpan.TicksPerSecond * 10)
        Console.WriteLine(ts.TotalSeconds)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}

#[test]
fn test_vb_time_span_addition_operator() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim t1 = TimeSpan.FromHours(2)
        Dim t2 = TimeSpan.FromMinutes(30)
        Dim sum = t1 + t2
        Console.WriteLine(sum.TotalMinutes)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["150"]);
}

#[test]
fn test_vb_time_span_subtraction_operator() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim t1 = TimeSpan.FromHours(5)
        Dim t2 = TimeSpan.FromHours(2)
        Dim diff = t1 - t2
        Console.WriteLine(diff.TotalHours)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_time_span_negation_operator() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ts = TimeSpan.FromHours(3)
        Dim neg = -ts
        Console.WriteLine(neg.TotalHours)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-3"]);
}

#[test]
fn test_vb_time_span_multiplication_operator() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ts = TimeSpan.FromMinutes(15)
        Dim scaled = ts * 4
        Console.WriteLine(scaled.TotalHours)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_time_span_division_operator() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ts = TimeSpan.FromHours(4)
        Dim half = ts / 2
        Console.WriteLine(half.TotalHours)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_time_span_duration_absolute_value() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim neg = TimeSpan.FromHours(-5)
        Dim abs = neg.Duration()
        Console.WriteLine(abs.TotalHours)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5"]);
}

#[test]
fn test_vb_time_span_comparison_operators() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim t1 = TimeSpan.FromMinutes(10)
        Dim t2 = TimeSpan.FromMinutes(20)
        Console.WriteLine((t1 < t2) & "|" & (t1 = t2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_time_span_zero_and_min_max_values() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine((TimeSpan.Zero.Ticks = 0L) & "|" & (TimeSpan.MaxValue > TimeSpan.MinValue))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_time_span_parse_standard_format() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ts = TimeSpan.Parse("02:30:45")
        Console.WriteLine(ts.Hours & ":" & ts.Minutes & ":" & ts.Seconds)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2:30:45"]);
}

#[test]
fn test_vb_time_span_try_parse_success_and_failure() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ts As TimeSpan
        Dim ok = TimeSpan.TryParse("1.12:00:00", ts)
        Dim fail = TimeSpan.TryParse("Invalid", ts)
        Console.WriteLine(ok & ":" & ts.Days & "|" & fail)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True:1|False"]);
}

#[test]
fn test_vb_time_span_to_string_custom_format() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ts = TimeSpan.FromMinutes(125)
        Console.WriteLine(ts.ToString("hh\:mm"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["02:05"]);
}

#[test]
fn test_vb_time_span_constructor_h_m_s() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ts As New TimeSpan(3, 45, 30)
        Console.WriteLine(ts.TotalSeconds)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["13530"]);
}

#[test]
fn test_vb_time_span_constructor_d_h_m_s_ms() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ts As New TimeSpan(1, 2, 3, 4, 500)
        Console.WriteLine(ts.Days & "d " & ts.Hours & "h " & ts.Minutes & "m " & ts.Seconds & "s " & ts.Milliseconds & "ms")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1d 2h 3m 4s 500ms"]);
}

#[test]
fn test_vb_time_span_equality_hashcode() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim t1 = TimeSpan.FromMinutes(60)
        Dim t2 = TimeSpan.FromHours(1)
        Console.WriteLine((t1 = t2) & "|" & (t1.GetHashCode() = t2.GetHashCode()))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}
