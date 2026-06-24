use crate::helpers::run_main;

#[test]
fn local_date_of_reads_year_month_day() {
    let out = run_main(
        r#"java.time.LocalDate d = java.time.LocalDate.of(2024, 6, 15); System.out.println(d.getYear()); System.out.println(d.getMonthValue()); System.out.println(d.getDayOfMonth());"#,
    );
    assert_eq!(out, vec!["2024", "6", "15"]);
}

#[test]
fn local_date_parse_iso_date_string() {
    let out = run_main(
        r#"java.time.LocalDate d = java.time.LocalDate.parse("2024-06-15"); System.out.println(d.getYear()); System.out.println(d.getMonthValue());"#,
    );
    assert_eq!(out, vec!["2024", "6"]);
}

#[test]
fn local_date_plus_days_advances_calendar() {
    let out = run_main(
        r#"java.time.LocalDate d = java.time.LocalDate.of(2024, 1, 30); System.out.println(d.plusDays(5).getDayOfMonth());"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn local_date_minus_days_moves_backward() {
    let out = run_main(
        r#"java.time.LocalDate d = java.time.LocalDate.of(2024, 3, 10); System.out.println(d.minusDays(3).getDayOfMonth());"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn local_date_plus_months_wraps_year_boundary() {
    let out = run_main(
        r#"java.time.LocalDate d = java.time.LocalDate.of(2024, 11, 15); System.out.println(d.plusMonths(2).getMonthValue());"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn local_date_minus_months_moves_to_prior_month() {
    let out = run_main(
        r#"java.time.LocalDate d = java.time.LocalDate.of(2024, 3, 1); System.out.println(d.minusMonths(1).getMonthValue());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn local_date_compare_to_equal_dates_is_zero() {
    let out = run_main(
        r#"java.time.LocalDate a = java.time.LocalDate.of(2024, 5, 1); java.time.LocalDate b = java.time.LocalDate.of(2024, 5, 1); System.out.println(a.compareTo(b));"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn local_date_compare_to_earlier_date_is_negative() {
    let out = run_main(
        r#"java.time.LocalDate a = java.time.LocalDate.of(2024, 1, 1); java.time.LocalDate b = java.time.LocalDate.of(2024, 6, 1); System.out.println(a.compareTo(b) < 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn local_date_compare_to_later_date_is_positive() {
    let out = run_main(
        r#"java.time.LocalDate a = java.time.LocalDate.of(2024, 12, 1); java.time.LocalDate b = java.time.LocalDate.of(2024, 1, 1); System.out.println(a.compareTo(b) > 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn local_date_is_before_detects_earlier_date() {
    let out = run_main(
        r#"java.time.LocalDate a = java.time.LocalDate.of(2024, 1, 1); java.time.LocalDate b = java.time.LocalDate.of(2024, 2, 1); System.out.println(a.isBefore(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn local_date_is_after_detects_later_date() {
    let out = run_main(
        r#"java.time.LocalDate a = java.time.LocalDate.of(2024, 3, 1); java.time.LocalDate b = java.time.LocalDate.of(2024, 2, 1); System.out.println(a.isAfter(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn local_date_is_equal_for_same_calendar_day() {
    let out = run_main(
        r#"java.time.LocalDate a = java.time.LocalDate.parse("2024-07-04"); java.time.LocalDate b = java.time.LocalDate.parse("2024-07-04"); System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn local_date_to_string_uses_iso_format() {
    let out = run_main(
        r#"java.time.LocalDate d = java.time.LocalDate.of(2024, 6, 15); System.out.println(d.toString());"#,
    );
    assert_eq!(out, vec!["2024-06-15"]);
}

#[test]
fn local_date_length_of_month_february_leap_year() {
    let out = run_main(
        r#"java.time.LocalDate d = java.time.LocalDate.of(2024, 2, 1); System.out.println(d.lengthOfMonth());"#,
    );
    assert_eq!(out, vec!["29"]);
}

#[test]
fn local_date_is_leap_year_for_2024() {
    let out = run_main(
        r#"java.time.LocalDate d = java.time.LocalDate.of(2024, 1, 1); System.out.println(d.isLeapYear());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn local_time_of_reads_hour_minute_second() {
    let out = run_main(
        r#"java.time.LocalTime t = java.time.LocalTime.of(14, 30, 45); System.out.println(t.getHour()); System.out.println(t.getMinute()); System.out.println(t.getSecond());"#,
    );
    assert_eq!(out, vec!["14", "30", "45"]);
}

#[test]
fn local_time_parse_iso_time_string() {
    let out = run_main(
        r#"java.time.LocalTime t = java.time.LocalTime.parse("09:15:00"); System.out.println(t.getHour()); System.out.println(t.getMinute());"#,
    );
    assert_eq!(out, vec!["9", "15"]);
}

#[test]
fn local_time_plus_hours_advances_clock() {
    let out = run_main(
        r#"java.time.LocalTime t = java.time.LocalTime.of(10, 0); System.out.println(t.plusHours(3).getHour());"#,
    );
    assert_eq!(out, vec!["13"]);
}

#[test]
fn local_time_minus_hours_moves_backward() {
    let out = run_main(
        r#"java.time.LocalTime t = java.time.LocalTime.of(5, 0); System.out.println(t.minusHours(2).getHour());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn local_time_plus_minutes_increments_minute_field() {
    let out = run_main(
        r#"java.time.LocalTime t = java.time.LocalTime.of(12, 50); System.out.println(t.plusMinutes(15).getMinute());"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn local_time_minus_seconds_decrements_second_field() {
    let out = run_main(
        r#"java.time.LocalTime t = java.time.LocalTime.of(0, 0, 30); System.out.println(t.minusSeconds(10).getSecond());"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn local_time_compare_to_equal_times_is_zero() {
    let out = run_main(
        r#"java.time.LocalTime a = java.time.LocalTime.of(8, 0); java.time.LocalTime b = java.time.LocalTime.of(8, 0); System.out.println(a.compareTo(b));"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn local_time_is_before_detects_earlier_time() {
    let out = run_main(
        r#"java.time.LocalTime a = java.time.LocalTime.of(8, 0); java.time.LocalTime b = java.time.LocalTime.of(9, 0); System.out.println(a.isBefore(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn local_time_is_after_detects_later_time() {
    let out = run_main(
        r#"java.time.LocalTime a = java.time.LocalTime.of(18, 0); java.time.LocalTime b = java.time.LocalTime.of(17, 0); System.out.println(a.isAfter(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn local_date_time_of_combines_date_and_time() {
    let out = run_main(
        r#"java.time.LocalDateTime dt = java.time.LocalDateTime.of(2024, 6, 15, 10, 30); System.out.println(dt.getYear()); System.out.println(dt.getHour());"#,
    );
    assert_eq!(out, vec!["2024", "10"]);
}

#[test]
fn local_date_time_parse_iso_date_time_string() {
    let out = run_main(
        r#"java.time.LocalDateTime dt = java.time.LocalDateTime.parse("2024-06-15T10:30:00"); System.out.println(dt.getDayOfMonth()); System.out.println(dt.getMinute());"#,
    );
    assert_eq!(out, vec!["15", "30"]);
}

#[test]
fn local_date_time_plus_days_advances_date_component() {
    let out = run_main(
        r#"java.time.LocalDateTime dt = java.time.LocalDateTime.of(2024, 1, 31, 12, 0); System.out.println(dt.plusDays(1).getDayOfMonth());"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn local_date_time_minus_hours_adjusts_time_component() {
    let out = run_main(
        r#"java.time.LocalDateTime dt = java.time.LocalDateTime.of(2024, 6, 1, 2, 0); System.out.println(dt.minusHours(3).getHour());"#,
    );
    assert_eq!(out, vec!["23"]);
}

#[test]
fn local_date_time_to_local_date_strips_time() {
    let out = run_main(
        r#"java.time.LocalDateTime dt = java.time.LocalDateTime.parse("2024-06-15T10:30:00"); System.out.println(dt.toLocalDate().toString());"#,
    );
    assert_eq!(out, vec!["2024-06-15"]);
}

#[test]
fn local_date_time_to_local_time_strips_date() {
    let out = run_main(
        r#"java.time.LocalDateTime dt = java.time.LocalDateTime.parse("2024-06-15T10:30:00"); System.out.println(dt.toLocalTime().getHour());"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn period_of_days_reads_day_count() {
    let out = run_main(
        r#"java.time.Period p = java.time.Period.ofDays(10); System.out.println(p.getDays());"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn period_of_months_reads_month_count() {
    let out = run_main(
        r#"java.time.Period p = java.time.Period.ofMonths(3); System.out.println(p.getMonths());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn period_between_two_dates_counts_days() {
    let out = run_main(
        r#"java.time.LocalDate a = java.time.LocalDate.of(2024, 1, 1); java.time.LocalDate b = java.time.LocalDate.of(2024, 1, 8); System.out.println(java.time.Period.between(a, b).getDays());"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn duration_of_hours_reads_hour_count() {
    let out = run_main(
        r#"java.time.Duration d = java.time.Duration.ofHours(2); System.out.println(d.toHours());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn duration_of_minutes_reads_minute_count() {
    let out = run_main(
        r#"java.time.Duration d = java.time.Duration.ofMinutes(90); System.out.println(d.toMinutes());"#,
    );
    assert_eq!(out, vec!["90"]);
}

#[test]
fn duration_of_seconds_reads_second_count() {
    let out = run_main(
        r#"java.time.Duration d = java.time.Duration.ofSeconds(45); System.out.println(d.getSeconds());"#,
    );
    assert_eq!(out, vec!["45"]);
}

#[test]
fn duration_between_two_times_counts_minutes() {
    let out = run_main(
        r#"java.time.LocalTime a = java.time.LocalTime.of(10, 0); java.time.LocalTime b = java.time.LocalTime.of(10, 30); System.out.println(java.time.Duration.between(a, b).toMinutes());"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn duration_plus_hours_adds_to_existing_duration() {
    let out = run_main(
        r#"java.time.Duration d = java.time.Duration.ofHours(1); System.out.println(d.plusHours(2).toHours());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn duration_minus_minutes_subtracts_from_duration() {
    let out = run_main(
        r#"java.time.Duration d = java.time.Duration.ofMinutes(60); System.out.println(d.minusMinutes(15).toMinutes());"#,
    );
    assert_eq!(out, vec!["45"]);
}

#[test]
fn local_date_time_compare_to_orders_chronologically() {
    let out = run_main(
        r#"java.time.LocalDateTime early = java.time.LocalDateTime.parse("2024-01-01T00:00:00"); java.time.LocalDateTime late = java.time.LocalDateTime.parse("2024-12-31T23:59:59"); System.out.println(early.isBefore(late)); System.out.println(late.isAfter(early));"#,
    );
    assert_eq!(out, vec!["true", "true"]);
}
