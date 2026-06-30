use crate::helpers::run_main;

#[test]
fn offset_date_time_of_reads_year() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 6, 15, 10, 30, 0, 0, java.time.ZoneOffset.UTC); System.out.println(o.getYear());"#);
    assert_eq!(out, vec!["2024"]);
}

#[test]
fn offset_date_time_of_reads_month() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 6, 15, 10, 30, 0, 0, java.time.ZoneOffset.UTC); System.out.println(o.getMonthValue());"#);
    assert_eq!(out, vec!["6"]);
}

#[test]
fn offset_date_time_of_reads_day() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 6, 15, 10, 30, 0, 0, java.time.ZoneOffset.UTC); System.out.println(o.getDayOfMonth());"#);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn offset_date_time_of_reads_hour() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 6, 15, 10, 30, 0, 0, java.time.ZoneOffset.UTC); System.out.println(o.getHour());"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn offset_date_time_of_reads_offset_seconds() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 1, 1, 0, 0, 0, 0, java.time.ZoneOffset.ofHours(2)); System.out.println(o.getOffset().getTotalSeconds());"#);
    assert_eq!(out, vec!["7200"]);
}

#[test]
fn offset_date_time_parse_iso() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T10:30:00+02:00"); System.out.println(o.getHour()); System.out.println(o.getOffset().getTotalSeconds() / 3600);"#);
    assert_eq!(out, vec!["10", "2"]);
}

#[test]
fn offset_date_time_parse_z_suffix() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T10:30:00Z"); System.out.println(o.getOffset().getTotalSeconds());"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn offset_date_time_plus_days() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 1, 31, 12, 0, 0, 0, java.time.ZoneOffset.UTC); System.out.println(o.plusDays(1).getDayOfMonth());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn offset_date_time_minus_hours() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 6, 1, 2, 0, 0, 0, java.time.ZoneOffset.UTC); System.out.println(o.minusHours(3).getHour());"#);
    assert_eq!(out, vec!["23"]);
}

#[test]
fn offset_date_time_plus_minutes() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 6, 1, 10, 50, 0, 0, java.time.ZoneOffset.UTC); System.out.println(o.plusMinutes(15).getMinute());"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn offset_date_time_to_local_date() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T10:30:00Z"); System.out.println(o.toLocalDate().toString());"#);
    assert_eq!(out, vec!["2024-06-15"]);
}

#[test]
fn offset_date_time_to_local_time() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T10:30:00Z"); System.out.println(o.toLocalTime().getMinute());"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn offset_date_time_to_local_date_time() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T10:30:00+01:00"); System.out.println(o.toLocalDateTime().getHour());"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn offset_date_time_to_instant() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("1970-01-01T00:00:00Z"); System.out.println(o.toInstant().getEpochSecond());"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn offset_date_time_to_zoned_date_time() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T10:30:00Z"); System.out.println(o.toZonedDateTime().getZone().getId());"#);
    assert_eq!(out, vec!["Z"]);
}

#[test]
fn offset_date_time_with_offset_same_instant() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T10:00:00+02:00"); System.out.println(o.withOffsetSameInstant(java.time.ZoneOffset.ofHours(4)).getHour());"#);
    assert_eq!(out, vec!["12"]);
}

#[test]
fn offset_date_time_with_offset_same_local() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T10:00:00+02:00"); System.out.println(o.withOffsetSameLocal(java.time.ZoneOffset.ofHours(4)).getHour());"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn offset_date_time_compare_to_equal() {
    let out = run_main(r#"java.time.OffsetDateTime a = java.time.OffsetDateTime.parse("2024-01-01T00:00:00Z"); java.time.OffsetDateTime b = java.time.OffsetDateTime.parse("2024-01-01T00:00:00Z"); System.out.println(a.compareTo(b));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn offset_date_time_is_before() {
    let out = run_main(r#"java.time.OffsetDateTime a = java.time.OffsetDateTime.parse("2024-01-01T00:00:00Z"); java.time.OffsetDateTime b = java.time.OffsetDateTime.parse("2024-02-01T00:00:00Z"); System.out.println(a.isBefore(b));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn offset_date_time_is_after() {
    let out = run_main(r#"java.time.OffsetDateTime a = java.time.OffsetDateTime.parse("2024-03-01T00:00:00Z"); java.time.OffsetDateTime b = java.time.OffsetDateTime.parse("2024-02-01T00:00:00Z"); System.out.println(a.isAfter(b));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn offset_date_time_equals_same() {
    let out = run_main(r#"java.time.OffsetDateTime a = java.time.OffsetDateTime.parse("2024-06-15T10:30:00Z"); java.time.OffsetDateTime b = java.time.OffsetDateTime.parse("2024-06-15T10:30:00Z"); System.out.println(a.equals(b));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn offset_date_time_to_string() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 6, 15, 10, 30, 0, 0, java.time.ZoneOffset.UTC); System.out.println(o.toString().contains("2024-06-15"));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn offset_date_time_at_zone_same_instant() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T12:00:00+02:00"); System.out.println(o.atZoneSameInstant(java.time.ZoneId.of("UTC")).getHour());"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn offset_date_time_at_zone_similar_local() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T12:00:00+02:00"); System.out.println(o.atZoneSimilarLocal(java.time.ZoneId.of("UTC")).getHour());"#);
    assert_eq!(out, vec!["12"]);
}

#[test]
fn offset_date_time_get_day_of_week() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 6, 17, 0, 0, 0, 0, java.time.ZoneOffset.UTC); System.out.println(o.getDayOfWeek().toString());"#);
    assert_eq!(out, vec!["MONDAY"]);
}

#[test]
fn offset_date_time_get_day_of_year() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 12, 31, 0, 0, 0, 0, java.time.ZoneOffset.UTC); System.out.println(o.getDayOfYear());"#);
    assert_eq!(out, vec!["366"]);
}

#[test]
fn offset_date_time_truncated_to_hours() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T10:30:45Z"); System.out.println(o.truncatedTo(java.time.temporal.ChronoUnit.HOURS).getMinute());"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn offset_date_time_with_year() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2023-06-15T10:00:00Z"); System.out.println(o.withYear(2025).getYear());"#);
    assert_eq!(out, vec!["2025"]);
}

#[test]
fn offset_date_time_with_month() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-01-15T10:00:00Z"); System.out.println(o.withMonth(12).getMonthValue());"#);
    assert_eq!(out, vec!["12"]);
}

#[test]
fn offset_date_time_with_day_of_month() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T10:00:00Z"); System.out.println(o.withDayOfMonth(1).getDayOfMonth());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn offset_date_time_with_hour() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T10:00:00Z"); System.out.println(o.withHour(18).getHour());"#);
    assert_eq!(out, vec!["18"]);
}

#[test]
fn offset_date_time_with_minute() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T10:30:00Z"); System.out.println(o.withMinute(0).getMinute());"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn offset_date_time_with_second() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T10:30:45Z"); System.out.println(o.withSecond(0).getSecond());"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn offset_date_time_of_local_date_time() {
    let out = run_main(r#"java.time.LocalDateTime ldt = java.time.LocalDateTime.of(2024, 3, 1, 8, 0); java.time.OffsetDateTime o = java.time.OffsetDateTime.of(ldt, java.time.ZoneOffset.ofHours(1)); System.out.println(o.getHour()); System.out.println(o.getOffset().getTotalSeconds() / 3600);"#);
    assert_eq!(out, vec!["8", "1"]);
}

#[test]
fn offset_date_time_of_instant_and_offset() {
    let out = run_main(r#"java.time.Instant i = java.time.Instant.ofEpochSecond(3600); java.time.OffsetDateTime o = java.time.OffsetDateTime.ofInstant(i, java.time.ZoneOffset.UTC); System.out.println(o.getHour());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn offset_date_time_minus_days() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 3, 10, 0, 0, 0, 0, java.time.ZoneOffset.UTC); System.out.println(o.minusDays(3).getDayOfMonth());"#);
    assert_eq!(out, vec!["7"]);
}

#[test]
fn offset_date_time_plus_months() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 11, 15, 0, 0, 0, 0, java.time.ZoneOffset.UTC); System.out.println(o.plusMonths(2).getMonthValue());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn offset_date_time_minus_months() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 3, 1, 0, 0, 0, 0, java.time.ZoneOffset.UTC); System.out.println(o.minusMonths(1).getMonthValue());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn offset_date_time_plus_seconds() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T10:00:00Z"); System.out.println(o.plusSeconds(90).getMinute());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn offset_date_time_minus_seconds() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.parse("2024-06-15T10:02:00Z"); System.out.println(o.minusSeconds(60).getMinute());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn offset_date_time_hash_code_equal() {
    let out = run_main(r#"java.time.OffsetDateTime a = java.time.OffsetDateTime.parse("2024-06-15T10:00:00Z"); java.time.OffsetDateTime b = java.time.OffsetDateTime.parse("2024-06-15T10:00:00Z"); System.out.println(a.hashCode() == b.hashCode());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn offset_date_time_get_nano() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 1, 1, 0, 0, 0, 123456789, java.time.ZoneOffset.UTC); System.out.println(o.getNano());"#);
    assert_eq!(out, vec!["123456789"]);
}

#[test]
fn offset_date_time_is_equal_same_values() {
    let out = run_main(r#"java.time.OffsetDateTime a = java.time.OffsetDateTime.parse("2024-06-15T10:00:00+01:00"); java.time.OffsetDateTime b = java.time.OffsetDateTime.parse("2024-06-15T10:00:00+01:00"); System.out.println(a.isEqual(b));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn offset_date_time_negative_offset() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 6, 15, 10, 0, 0, 0, java.time.ZoneOffset.ofHours(-7)); System.out.println(o.getOffset().getTotalSeconds() / 3600);"#);
    assert_eq!(out, vec!["-7"]);
}

#[test]
fn offset_date_time_format_iso() {
    let out = run_main(r#"java.time.OffsetDateTime o = java.time.OffsetDateTime.of(2024, 6, 15, 10, 30, 0, 0, java.time.ZoneOffset.UTC); System.out.println(o.format(java.time.format.DateTimeFormatter.ISO_OFFSET_DATE_TIME).contains("2024-06-15"));"#);
    assert_eq!(out, vec!["true"]);
}

