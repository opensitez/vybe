use super::helpers::run_vb;

#[test]
fn datetime_offset_from_components() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim dto As New DateTimeOffset(2024, 1, 2, 3, 4, 5, TimeSpan.FromHours(-5))
        Console.WriteLine(dto.Year)
        Console.WriteLine(dto.Month)
        Console.WriteLine(dto.Day)
        Console.WriteLine(dto.Offset.Hours)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2024", "1", "2", "-5"]);
}

#[test]
fn datetime_offset_utc_now_is_zero_offset() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim dto As DateTimeOffset = DateTimeOffset.UtcNow
        Console.WriteLine(dto.Offset = TimeSpan.Zero)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn datetime_offset_to_local_and_to_offset() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim source As New DateTimeOffset(2024, 1, 2, 3, 4, 5, TimeSpan.FromHours(2))
        Dim utc As DateTimeOffset = source.ToUniversalTime()
        Dim restored As DateTimeOffset = source.ToOffset(TimeSpan.FromHours(2))
        Console.WriteLine(utc.Offset = TimeSpan.Zero)
        Console.WriteLine(restored.Offset = TimeSpan.FromHours(2))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn datetime_offset_add_and_subtract() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim base As New DateTimeOffset(2024, 1, 1, 0, 0, 0, TimeSpan.Zero)
        Dim later As DateTimeOffset = base.AddDays(1).AddHours(3)
        Dim earlier As DateTimeOffset = later.AddHours(-3)
        Console.WriteLine(later > base)
        Console.WriteLine(earlier = base)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn datetime_offset_compare_to_operator() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim a As New DateTimeOffset(2024, 1, 1, 0, 0, 0, TimeSpan.FromHours(1))
        Dim b As New DateTimeOffset(2024, 1, 1, 1, 0, 0, TimeSpan.Zero)
        Console.WriteLine(a = b)
        Console.WriteLine(a.CompareTo(b))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "0"]);
}

#[test]
fn datetime_offset_to_unix_time_is_stable() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim dto As New DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero)
        Console.WriteLine(dto.ToUnixTimeSeconds())
        Dim one As DateTimeOffset = DateTimeOffset.FromUnixTimeSeconds(0)
        Console.WriteLine(one = dto)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["0", "True"]);
}

#[test]
fn datetime_offset_parse_roundtrip() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim dto As DateTimeOffset = DateTimeOffset.Parse("2024-07-21T03:04:05+02:00")
        Dim text As String = dto.ToString("o")
        Dim again As DateTimeOffset = DateTimeOffset.Parse(text)
        Console.WriteLine(dto = again)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}
