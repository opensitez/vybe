//! `DateTimeOffset`, `TimeZoneInfo`, and cross-timezone arithmetic.
use super::helpers::run_csharp;

#[test]
fn datetimeoffset_stores_offset_component() {
    assert_eq!(
        run_csharp(
            r#"var dto=new System.DateTimeOffset(2024,1,15,10,0,0,System.TimeSpan.FromHours(5));
Console.WriteLine(dto.Offset.Hours);"#
        ),
        &["5"]
    );
}

#[test]
fn datetimeoffset_utc_has_zero_offset() {
    assert_eq!(
        run_csharp(
            r#"var dto=System.DateTimeOffset.UtcNow;
Console.WriteLine(dto.Offset==System.TimeSpan.Zero);"#
        ),
        &["True"]
    );
}

#[test]
fn datetimeoffset_to_universal_time_yields_utc() {
    assert_eq!(
        run_csharp(
            r#"var dto=new System.DateTimeOffset(2024,1,15,10,0,0,System.TimeSpan.FromHours(2));
var utc=dto.ToUniversalTime();
Console.WriteLine(utc.Hour);"#
        ),
        &["8"]
    );
}

#[test]
fn datetime_to_universal_time_converts_to_utc_kind() {
    assert_eq!(
        run_csharp(
            r#"var local=new System.DateTime(2024,1,15,12,0,0,System.DateTimeKind.Local);
var utc=local.ToUniversalTime();
Console.WriteLine(utc.Kind);"#
        ),
        &["Utc"]
    );
}

#[test]
fn timespan_negate_inverts_direction() {
    assert_eq!(
        run_csharp(
            r#"var ts=System.TimeSpan.FromHours(3);
Console.WriteLine((-ts).Hours);"#
        ),
        &["-3"]
    );
}

#[test]
fn datetimeoffset_add_hours_adjusts_wall_time() {
    assert_eq!(
        run_csharp(
            r#"var dto=new System.DateTimeOffset(2024,1,1,20,0,0,System.TimeSpan.Zero);
var next=dto.AddHours(5);
Console.WriteLine(next.Day); Console.WriteLine(next.Hour);"#
        ),
        &["2", "1"]
    );
}
