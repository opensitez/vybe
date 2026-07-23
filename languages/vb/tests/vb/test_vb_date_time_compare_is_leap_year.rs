use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: DateTime.Compare, IsLeapYear & Date Comparisons
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_date_time_compare_earlier_later_equal() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim d1 As New DateTime(2025, 1, 1)
        Dim d2 As New DateTime(2025, 1, 2)
        Console.WriteLine(DateTime.Compare(d1, d2) < 0 & "|" & DateTime.Compare(d2, d1) > 0 & "|" & DateTime.Compare(d1, d1) = 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|True"]);
}

#[test]
fn test_vb_date_time_compare_to_instance_method() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim d1 As New DateTime(2025, 6, 1)
        Dim d2 As New DateTime(2025, 6, 1)
        Console.WriteLine(d1.CompareTo(d2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_date_time_is_leap_year_century_rule() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' 2000 was leap year (divisible by 400), 1900 was NOT leap year (divisible by 100 but not 400), 2024 was leap year
        Console.WriteLine(DateTime.IsLeapYear(2000) & "|" & DateTime.IsLeapYear(1900) & "|" & DateTime.IsLeapYear(2024))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False|True"]);
}

#[test]
fn test_vb_date_time_equals_static_method() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim d1 As New DateTime(2025, 1, 1, 10, 0, 0)
        Dim d2 As New DateTime(2025, 1, 1, 10, 0, 0)
        Dim d3 As New DateTime(2025, 1, 1, 11, 0, 0)
        Console.WriteLine(DateTime.Equals(d1, d2) & "|" & DateTime.Equals(d1, d3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_date_time_operators_relational() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim d1 As New DateTime(2025, 1, 1)
        Dim d2 As New DateTime(2025, 1, 2)
        Console.WriteLine((d1 < d2) & "|" & (d1 <= d2) & "|" & (d2 > d1) & "|" & (d2 >= d1) & "|" & (d1 <> d2))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True|True|True|True"]);
}

#[test]
fn test_vb_date_time_array_sort() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dates As DateTime() = {New DateTime(2025, 3, 1), New DateTime(2025, 1, 1), New DateTime(2025, 2, 1)}
        Array.Sort(dates)
        For Each d In dates
            Console.WriteLine(d.Month)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1", "2", "3"]);
}

#[test]
fn test_vb_date_time_is_daylight_saving_time_simulation() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dtUtc As New DateTime(2025, 6, 1, 0, 0, 0, DateTimeKind.Utc)
        Console.WriteLine(dtUtc.Kind.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Utc"]);
}

#[test]
fn test_vb_date_time_between_range_check() {
    let src = r#"
Imports System

Module Program
    Private Function IsInRange(target As DateTime, startDt As DateTime, endDt As DateTime) As Boolean
        Return target >= startDt AndAlso target <= endDt
    End Function

    Sub Main()
        Dim startDt As New DateTime(2025, 1, 1)
        Dim endDt As New DateTime(2025, 12, 31)
        Dim testDt As New DateTime(2025, 6, 15)
        Console.WriteLine(IsInRange(testDt, startDt, endDt))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_date_time_compare_date_component_only() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim d1 As New DateTime(2025, 5, 10, 8, 0, 0)
        Dim d2 As New DateTime(2025, 5, 10, 20, 0, 0)
        Console.WriteLine(d1.Date = d2.Date)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_date_time_min_max_static_helpers() {
    let src = r#"
Imports System

Module Program
    Private Function MaxDate(d1 As DateTime, d2 As DateTime) As DateTime
        If d1 > d2 Then Return d1 Else Return d2
    End Function

    Sub Main()
        Dim d1 As New DateTime(2025, 1, 1)
        Dim d2 As New DateTime(2025, 5, 1)
        Console.WriteLine(MaxDate(d1, d2).Month)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5"]);
}

#[test]
fn test_vb_date_time_is_leap_year_negative_year_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            DateTime.IsLeapYear(0)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("ArgumentOutOfRangeException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ArgumentOutOfRangeException Caught"]);
}

#[test]
fn test_vb_date_time_compare_ticks_direct() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim d1 As New DateTime(2025, 1, 1)
        Dim d2 As New DateTime(2025, 1, 1, 0, 0, 1)
        Console.WriteLine(d1.Ticks < d2.Ticks)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_date_time_binary_representation_to_from() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim d1 As New DateTime(2025, 10, 31, 15, 45, 0, DateTimeKind.Utc)
        Dim bin = d1.ToBinary()
        Dim d2 = DateTime.FromBinary(bin)
        Console.WriteLine((d1 = d2) & "|" & d2.Kind.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|Utc"]);
}

#[test]
fn test_vb_date_time_from_file_time_utc() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 1, 1, 0, 0, 0, DateTimeKind.Utc)
        Dim ft = dt.ToFileTimeUtc()
        Dim dtRestored = DateTime.FromFileTimeUtc(ft)
        Console.WriteLine(dt = dtRestored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_date_time_hash_code_equality() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim d1 As New DateTime(2025, 7, 7, 12, 0, 0)
        Dim d2 As New DateTime(2025, 7, 7, 12, 0, 0)
        Console.WriteLine(d1.GetHashCode() = d2.GetHashCode())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_date_time_to_oadate_from_oadate() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dt As New DateTime(2025, 1, 1)
        Dim oa = dt.ToOADate()
        Dim dtRestored = DateTime.FromOADate(oa)
        Console.WriteLine(dt = dtRestored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_date_time_list_binary_search() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of DateTime) From {
            New DateTime(2025, 1, 1),
            New DateTime(2025, 2, 1),
            New DateTime(2025, 3, 1)
        }
        Dim idx = list.BinarySearch(New DateTime(2025, 2, 1))
        Console.WriteLine(idx)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_date_time_linq_order_by() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim dates = {New DateTime(2025, 12, 1), New DateTime(2025, 1, 1), New DateTime(2025, 6, 1)}
        Dim sorted = dates.OrderBy(Function(d) d)
        Console.WriteLine(String.Join(",", sorted.Select(Function(d) d.Month)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,6,12"]);
}

#[test]
fn test_vb_date_time_linq_max_min() {
    let src = r#"
Imports System.Linq

Module Program
    Sub Main()
        Dim dates = {New DateTime(2025, 12, 1), New DateTime(2025, 1, 1), New DateTime(2025, 6, 1)}
        Console.WriteLine(dates.Min().Month & "|" & dates.Max().Month)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1|12"]);
}

#[test]
fn test_vb_date_time_type_of_check() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim obj As Object = New DateTime(2025, 1, 1)
        Console.WriteLine(TypeOf obj Is DateTime)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
