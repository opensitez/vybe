use crate::helpers::run_main;

#[test]
fn zoned_date_time_of_reads_year() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 6, 15, 10, 30, 0, 0, java.time.ZoneId.of("UTC")); System.out.println(z.getYear());"#);
    assert_eq!(out, vec!["2024"]);
}

#[test]
fn zoned_date_time_of_reads_zone() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 6, 15, 10, 30, 0, 0, java.time.ZoneId.of("UTC")); System.out.println(z.getZone().getId());"#);
    assert_eq!(out, vec!["UTC"]);
}

#[test]
fn zoned_date_time_parse_iso_z() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2024-06-15T10:30:00Z"); System.out.println(z.getHour()); System.out.println(z.getZone().getId());"#);
    assert_eq!(out, vec!["10", "Z"]);
}

#[test]
fn zoned_date_time_parse_with_offset() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2024-06-15T10:30:00+02:00[Europe/Paris]"); System.out.println(z.getOffset().getTotalSeconds() / 3600);"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn zoned_date_time_of_local_date_time() {
    let out = run_main(r#"java.time.LocalDateTime ldt = java.time.LocalDateTime.of(2024, 3, 1, 8, 0); java.time.ZonedDateTime z = java.time.ZonedDateTime.of(ldt, java.time.ZoneId.of("UTC")); System.out.println(z.getDayOfMonth());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn zoned_date_time_of_instant() {
    let out = run_main(r#"java.time.Instant i = java.time.Instant.ofEpochSecond(0); java.time.ZonedDateTime z = java.time.ZonedDateTime.ofInstant(i, java.time.ZoneId.of("UTC")); System.out.println(z.getYear());"#);
    assert_eq!(out, vec!["1970"]);
}

#[test]
fn zoned_date_time_get_offset_utc() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2024-06-15T10:00:00Z"); System.out.println(z.getOffset().getTotalSeconds());"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn zoned_date_time_with_zone_same_instant() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2024-06-15T10:00:00+02:00[Europe/Paris]"); System.out.println(z.withZoneSameInstant(java.time.ZoneId.of("UTC")).getHour());"#);
    assert_eq!(out, vec!["8"]);
}

#[test]
fn zoned_date_time_with_zone_same_local() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2024-06-15T10:00:00+02:00[Europe/Paris]"); System.out.println(z.withZoneSameLocal(java.time.ZoneId.of("UTC")).getHour());"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn zoned_date_time_to_offset_date_time() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2024-06-15T10:30:00Z"); System.out.println(z.toOffsetDateTime().getMinute());"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn zoned_date_time_to_instant() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("1970-01-01T01:00:00Z"); System.out.println(z.toInstant().getEpochSecond());"#);
    assert_eq!(out, vec!["3600"]);
}

#[test]
fn zoned_date_time_to_local_date_time() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2024-06-15T10:30:00Z"); System.out.println(z.toLocalDateTime().getHour());"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn zoned_date_time_to_local_date() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2024-06-15T10:30:00Z"); System.out.println(z.toLocalDate().toString());"#);
    assert_eq!(out, vec!["2024-06-15"]);
}

#[test]
fn zoned_date_time_to_local_time() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2024-06-15T10:30:00Z"); System.out.println(z.toLocalTime().getSecond());"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn zoned_date_time_plus_days() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 1, 31, 12, 0, 0, 0, java.time.ZoneId.of("UTC")); System.out.println(z.plusDays(1).getDayOfMonth());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn zoned_date_time_minus_hours() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 6, 1, 2, 0, 0, 0, java.time.ZoneId.of("UTC")); System.out.println(z.minusHours(3).getHour());"#);
    assert_eq!(out, vec!["23"]);
}

#[test]
fn zoned_date_time_plus_months() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 11, 15, 0, 0, 0, 0, java.time.ZoneId.of("UTC")); System.out.println(z.plusMonths(2).getMonthValue());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn zoned_date_time_minus_days() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 3, 10, 0, 0, 0, 0, java.time.ZoneId.of("UTC")); System.out.println(z.minusDays(3).getDayOfMonth());"#);
    assert_eq!(out, vec!["7"]);
}

#[test]
fn zoned_date_time_compare_to_equal() {
    let out = run_main(r#"java.time.ZonedDateTime a = java.time.ZonedDateTime.parse("2024-01-01T00:00:00Z"); java.time.ZonedDateTime b = java.time.ZonedDateTime.parse("2024-01-01T00:00:00Z"); System.out.println(a.compareTo(b));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn zoned_date_time_is_before() {
    let out = run_main(r#"java.time.ZonedDateTime a = java.time.ZonedDateTime.parse("2024-01-01T00:00:00Z"); java.time.ZonedDateTime b = java.time.ZonedDateTime.parse("2024-02-01T00:00:00Z"); System.out.println(a.isBefore(b));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn zoned_date_time_is_after() {
    let out = run_main(r#"java.time.ZonedDateTime a = java.time.ZonedDateTime.parse("2024-03-01T00:00:00Z"); java.time.ZonedDateTime b = java.time.ZonedDateTime.parse("2024-02-01T00:00:00Z"); System.out.println(a.isAfter(b));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn zoned_date_time_equals_same() {
    let out = run_main(r#"java.time.ZonedDateTime a = java.time.ZonedDateTime.parse("2024-06-15T10:30:00Z"); java.time.ZonedDateTime b = java.time.ZonedDateTime.parse("2024-06-15T10:30:00Z"); System.out.println(a.equals(b));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn zoned_date_time_get_day_of_week() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 6, 17, 0, 0, 0, 0, java.time.ZoneId.of("UTC")); System.out.println(z.getDayOfWeek().toString());"#);
    assert_eq!(out, vec!["MONDAY"]);
}

#[test]
fn zoned_date_time_get_day_of_year() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 12, 31, 0, 0, 0, 0, java.time.ZoneId.of("UTC")); System.out.println(z.getDayOfYear());"#);
    assert_eq!(out, vec!["366"]);
}

#[test]
fn zoned_date_time_truncated_to_minutes() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2024-06-15T10:30:45Z"); System.out.println(z.truncatedTo(java.time.temporal.ChronoUnit.MINUTES).getSecond());"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn zoned_date_time_with_year() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2023-06-15T10:00:00Z"); System.out.println(z.withYear(2025).getYear());"#);
    assert_eq!(out, vec!["2025"]);
}

#[test]
fn zoned_date_time_with_month() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2024-01-15T10:00:00Z"); System.out.println(z.withMonth(12).getMonthValue());"#);
    assert_eq!(out, vec!["12"]);
}

#[test]
fn zoned_date_time_with_day_of_month() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2024-06-15T10:00:00Z"); System.out.println(z.withDayOfMonth(1).getDayOfMonth());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn zoned_date_time_with_hour() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2024-06-15T10:00:00Z"); System.out.println(z.withHour(18).getHour());"#);
    assert_eq!(out, vec!["18"]);
}

#[test]
fn zoned_date_time_with_fixed_offset_zone() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 6, 15, 10, 0, 0, 0, java.time.ZoneId.of("+03:00")); System.out.println(z.getOffset().getTotalSeconds() / 3600);"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn zoned_date_time_to_string_contains_date() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 6, 15, 10, 30, 0, 0, java.time.ZoneId.of("UTC")); System.out.println(z.toString().contains("2024-06-15"));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn zoned_date_time_hash_code_equal() {
    let out = run_main(r#"java.time.ZonedDateTime a = java.time.ZonedDateTime.parse("2024-06-15T10:00:00Z"); java.time.ZonedDateTime b = java.time.ZonedDateTime.parse("2024-06-15T10:00:00Z"); System.out.println(a.hashCode() == b.hashCode());"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn zoned_date_time_is_equal() {
    let out = run_main(r#"java.time.ZonedDateTime a = java.time.ZonedDateTime.parse("2024-06-15T10:00:00Z"); java.time.ZonedDateTime b = java.time.ZonedDateTime.parse("2024-06-15T10:00:00Z"); System.out.println(a.isEqual(b));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn zoned_date_time_plus_minutes() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 6, 1, 10, 50, 0, 0, java.time.ZoneId.of("UTC")); System.out.println(z.plusMinutes(15).getMinute());"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn zoned_date_time_minus_seconds() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2024-06-15T10:02:30Z"); System.out.println(z.minusSeconds(90).getMinute());"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn zoned_date_time_get_nano() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 1, 1, 0, 0, 0, 500000000, java.time.ZoneId.of("UTC")); System.out.println(z.getNano());"#);
    assert_eq!(out, vec!["500000000"]);
}

#[test]
fn zoned_date_time_format_iso() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 6, 15, 10, 30, 0, 0, java.time.ZoneId.of("UTC")); System.out.println(z.format(java.time.format.DateTimeFormatter.ISO_ZONED_DATE_TIME).contains("2024-06-15"));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn zoned_date_time_of_strict_handles_valid() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.ofStrict(java.time.LocalDateTime.of(2024, 6, 15, 10, 0), java.time.ZoneOffset.UTC, java.time.ZoneId.of("UTC")); System.out.println(z.getHour());"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn zoned_date_time_with_later_offset_at_overlap() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2018-10-28T02:30:00+02:00[Europe/Berlin]"); System.out.println(z.withLaterOffsetAtOverlap().getOffset().getTotalSeconds() / 3600);"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn zoned_date_time_with_earlier_offset_at_overlap() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2018-10-28T02:30:00+01:00[Europe/Berlin]"); System.out.println(z.withEarlierOffsetAtOverlap().getOffset().getTotalSeconds() / 3600);"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn zoned_date_time_to_epoch_second() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("1970-01-01T01:00:00Z"); System.out.println(z.toEpochSecond());"#);
    assert_eq!(out, vec!["3600"]);
}

#[test]
fn zoned_date_time_range_day_of_month() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 2, 1, 0, 0, 0, 0, java.time.ZoneId.of("UTC")); System.out.println(z.range(java.time.temporal.ChronoField.DAY_OF_MONTH).getMaximum());"#);
    assert_eq!(out, vec!["29"]);
}

#[test]
fn zoned_date_time_get_minute() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 6, 15, 10, 45, 0, 0, java.time.ZoneId.of("UTC")); System.out.println(z.getMinute());"#);
    assert_eq!(out, vec!["45"]);
}

#[test]
fn zoned_date_time_get_second() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 6, 15, 10, 30, 55, 0, java.time.ZoneId.of("UTC")); System.out.println(z.getSecond());"#);
    assert_eq!(out, vec!["55"]);
}

#[test]
fn zoned_date_time_with_zone_same_instant_to_tokyo() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.parse("2024-06-15T12:00:00Z"); System.out.println(z.withZoneSameInstant(java.time.ZoneId.of("Asia/Tokyo")).getHour());"#);
    assert_eq!(out, vec!["21"]);
}

#[test]
fn zoned_date_time_plus_weeks() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 6, 1, 0, 0, 0, 0, java.time.ZoneId.of("UTC")); System.out.println(z.plusWeeks(2).getDayOfMonth());"#);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn zoned_date_time_minus_months() {
    let out = run_main(r#"java.time.ZonedDateTime z = java.time.ZonedDateTime.of(2024, 3, 1, 0, 0, 0, 0, java.time.ZoneId.of("UTC")); System.out.println(z.minusMonths(1).getMonthValue());"#);
    assert_eq!(out, vec!["2"]);
}

