use crate::helpers::run_main;

#[test]
fn instant_epoch_second_zero() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(0); System.out.println(i.getEpochSecond());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn instant_epoch_second_positive() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(1700000000L); System.out.println(i.getEpochSecond());"#,
    );
    assert_eq!(out, vec!["1700000000"]);
}

#[test]
fn instant_epoch_second_with_nano() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(1, 500000000); System.out.println(i.getNano());"#,
    );
    assert_eq!(out, vec!["500000000"]);
}

#[test]
fn instant_of_epoch_milli() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochMilli(1500L); System.out.println(i.toEpochMilli());"#,
    );
    assert_eq!(out, vec!["1500"]);
}

#[test]
fn instant_parse_iso_string() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.parse("1970-01-01T00:00:00Z"); System.out.println(i.getEpochSecond());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn instant_parse_with_fraction() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.parse("1970-01-01T00:00:01.5Z"); System.out.println(i.getEpochSecond()); System.out.println(i.getNano());"#,
    );
    assert_eq!(out, vec!["1", "500000000"]);
}

#[test]
fn instant_plus_seconds() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(10); System.out.println(i.plusSeconds(5).getEpochSecond());"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn instant_minus_seconds() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(20); System.out.println(i.minusSeconds(7).getEpochSecond());"#,
    );
    assert_eq!(out, vec!["13"]);
}

#[test]
fn instant_plus_millis() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochMilli(1000L); System.out.println(i.plusMillis(250).toEpochMilli());"#,
    );
    assert_eq!(out, vec!["1250"]);
}

#[test]
fn instant_minus_millis() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochMilli(1000L); System.out.println(i.minusMillis(400).toEpochMilli());"#,
    );
    assert_eq!(out, vec!["600"]);
}

#[test]
fn instant_plus_nanos() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(0, 100); System.out.println(i.plusNanos(900).getNano());"#,
    );
    assert_eq!(out, vec!["1000"]);
}

#[test]
fn instant_minus_nanos() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(0, 1000); System.out.println(i.minusNanos(500).getNano());"#,
    );
    assert_eq!(out, vec!["500"]);
}

#[test]
fn instant_is_before_earlier() {
    let out = run_main(
        r#"java.time.Instant a = java.time.Instant.ofEpochSecond(1); java.time.Instant b = java.time.Instant.ofEpochSecond(2); System.out.println(a.isBefore(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instant_is_after_later() {
    let out = run_main(
        r#"java.time.Instant a = java.time.Instant.ofEpochSecond(5); java.time.Instant b = java.time.Instant.ofEpochSecond(3); System.out.println(a.isAfter(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instant_compare_to_equal() {
    let out = run_main(
        r#"java.time.Instant a = java.time.Instant.ofEpochSecond(42); java.time.Instant b = java.time.Instant.ofEpochSecond(42); System.out.println(a.compareTo(b));"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn instant_compare_to_negative() {
    let out = run_main(
        r#"java.time.Instant a = java.time.Instant.ofEpochSecond(1); java.time.Instant b = java.time.Instant.ofEpochSecond(9); System.out.println(a.compareTo(b) < 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instant_compare_to_positive() {
    let out = run_main(
        r#"java.time.Instant a = java.time.Instant.ofEpochSecond(9); java.time.Instant b = java.time.Instant.ofEpochSecond(1); System.out.println(a.compareTo(b) > 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instant_equals_same_epoch() {
    let out = run_main(
        r#"java.time.Instant a = java.time.Instant.ofEpochSecond(100, 50); java.time.Instant b = java.time.Instant.ofEpochSecond(100, 50); System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instant_equals_different_nano() {
    let out = run_main(
        r#"java.time.Instant a = java.time.Instant.ofEpochSecond(100, 50); java.time.Instant b = java.time.Instant.ofEpochSecond(100, 51); System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn instant_to_string_iso() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(0); System.out.println(i.toString());"#,
    );
    assert_eq!(out, vec!["1970-01-01T00:00:00Z"]);
}

#[test]
fn instant_at_zone_utc_year() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.parse("2024-06-15T10:30:00Z"); System.out.println(i.atZone(java.time.ZoneId.of("UTC")).getYear());"#,
    );
    assert_eq!(out, vec!["2024"]);
}

#[test]
fn instant_at_zone_utc_hour() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.parse("2024-06-15T10:30:00Z"); System.out.println(i.atZone(java.time.ZoneId.of("UTC")).getHour());"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn instant_at_offset_plus_two() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.parse("2024-06-15T08:00:00Z"); System.out.println(i.atOffset(java.time.ZoneOffset.ofHours(2)).getHour());"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn instant_truncated_to_seconds() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(5, 999999999); System.out.println(i.truncatedTo(java.time.temporal.ChronoUnit.SECONDS).getNano());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn instant_truncated_to_millis() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(0, 123456789); System.out.println(i.truncatedTo(java.time.temporal.ChronoUnit.MILLIS).getNano());"#,
    );
    assert_eq!(out, vec!["123000000"]);
}

#[test]
fn instant_between_seconds() {
    let out = run_main(
        r#"java.time.Instant a = java.time.Instant.ofEpochSecond(10); java.time.Instant b = java.time.Instant.ofEpochSecond(25); System.out.println(java.time.Duration.between(a, b).getSeconds());"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn instant_is_after_epoch() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(1); System.out.println(i.isAfter(java.time.Instant.EPOCH));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instant_is_before_future() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(1); System.out.println(i.isBefore(java.time.Instant.ofEpochSecond(100)));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instant_from_epoch_milli_zero() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochMilli(0L); System.out.println(i.getEpochSecond());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn instant_to_epoch_milli_roundtrip() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochMilli(2500L); System.out.println(i.toEpochMilli());"#,
    );
    assert_eq!(out, vec!["2500"]);
}

#[test]
fn instant_hash_code_consistent() {
    let out = run_main(
        r#"java.time.Instant a = java.time.Instant.ofEpochSecond(7); java.time.Instant b = java.time.Instant.ofEpochSecond(7); System.out.println(a.hashCode() == b.hashCode());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instant_at_zone_get_month() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.parse("2024-12-01T00:00:00Z"); System.out.println(i.atZone(java.time.ZoneId.of("UTC")).getMonthValue());"#,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn instant_at_zone_get_day() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.parse("2024-03-20T00:00:00Z"); System.out.println(i.atZone(java.time.ZoneId.of("UTC")).getDayOfMonth());"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn instant_plus_duration() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(100); System.out.println(i.plus(java.time.Duration.ofMinutes(2)).getEpochSecond());"#,
    );
    assert_eq!(out, vec!["220"]);
}

#[test]
fn instant_minus_duration() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(300); System.out.println(i.minus(java.time.Duration.ofMinutes(5)).getEpochSecond());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn instant_at_offset_get_offset_hours() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.parse("2024-01-01T12:00:00Z"); System.out.println(i.atOffset(java.time.ZoneOffset.ofHours(-5)).getOffset().getTotalSeconds() / 3600);"#,
    );
    assert_eq!(out, vec!["-5"]);
}

#[test]
fn instant_parse_leap_second_boundary() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.parse("2020-12-31T23:59:59Z"); System.out.println(i.getEpochSecond());"#,
    );
    assert_eq!(out, vec!["1609459199"]);
}

#[test]
fn instant_nano_zero_on_milli_only() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochMilli(5000L); System.out.println(i.getNano());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn instant_to_string_contains_t() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(3600); System.out.println(i.toString().contains("T"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instant_is_after_equal_false() {
    let out = run_main(
        r#"java.time.Instant a = java.time.Instant.ofEpochSecond(5); java.time.Instant b = java.time.Instant.ofEpochSecond(5); System.out.println(a.isAfter(b));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn instant_is_before_equal_false() {
    let out = run_main(
        r#"java.time.Instant a = java.time.Instant.ofEpochSecond(5); java.time.Instant b = java.time.Instant.ofEpochSecond(5); System.out.println(a.isBefore(b));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn instant_at_zone_to_local_date() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.parse("2024-07-04T00:00:00Z"); System.out.println(i.atZone(java.time.ZoneId.of("UTC")).toLocalDate().toString());"#,
    );
    assert_eq!(out, vec!["2024-07-04"]);
}

#[test]
fn instant_at_zone_to_local_time() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.parse("2024-07-04T15:45:30Z"); System.out.println(i.atZone(java.time.ZoneId.of("UTC")).toLocalTime().getMinute());"#,
    );
    assert_eq!(out, vec!["45"]);
}

#[test]
fn instant_of_epoch_second_negative() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(-86400L); System.out.println(i.getEpochSecond());"#,
    );
    assert_eq!(out, vec!["-86400"]);
}

#[test]
fn instant_plus_zero_seconds_unchanged() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(99); System.out.println(i.plusSeconds(0).getEpochSecond());"#,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn instant_minus_zero_nanos_unchanged() {
    let out = run_main(
        r#"java.time.Instant i = java.time.Instant.ofEpochSecond(0, 42); System.out.println(i.minusNanos(0).getNano());"#,
    );
    assert_eq!(out, vec!["42"]);
}
