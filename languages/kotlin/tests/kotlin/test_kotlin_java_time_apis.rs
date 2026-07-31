use crate::helpers::run_prints;

#[test]
fn test_local_date_parse_and_components() {
    let out = run_prints(r#"
        fun main() {
            val value = java.time.LocalDate.parse("2024-06-15")
            println(value.year)
            println(value.monthValue)
            println(value.dayOfMonth)
        }
    "#);
    assert_eq!(out, &["2024", "6", "15"]);
}

#[test]
fn test_local_date_is_leap_year() {
    let out = run_prints(r#"
        fun main() {
            println(java.time.LocalDate.parse("2024-02-01").isLeapYear())
            println(java.time.LocalDate.parse("2023-02-01").isLeapYear())
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_local_date_plus_days_is_deterministic() {
    let out = run_prints(r#"
        fun main() {
            val value = java.time.LocalDate.parse("2024-12-30").plusDays(5)
            println(value.toString())
        }
    "#);
    assert_eq!(out, &["2025-01-04"]);
}

#[test]
fn test_local_date_minus_months_and_days() {
    let out = run_prints(r#"
        fun main() {
            val value = java.time.LocalDate.parse("2024-03-31").minusMonths(1)
            println(value.toString())
            val shifted = value.minusDays(1)
            println(shifted.toString())
        }
    "#);
    assert_eq!(out, &["2024-02-29", "2024-02-28"]);
}

#[test]
fn test_local_date_compare_to() {
    let out = run_prints(r#"
        fun main() {
            val a = java.time.LocalDate.parse("2024-01-01")
            val b = java.time.LocalDate.parse("2024-01-02")
            println(a.isBefore(b))
            println(a.isAfter(b))
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_local_time_parse_and_components() {
    let out = run_prints(r#"
        fun main() {
            val value = java.time.LocalTime.parse("09:30:45")
            println(value.hour)
            println(value.minute)
            println(value.second)
        }
    "#);
    assert_eq!(out, &["9", "30", "45"]);
}

#[test]
fn test_local_time_plus_hours_minutes() {
    let out = run_prints(r#"
        fun main() {
            val value = java.time.LocalTime.parse("23:45:00").plusHours(2).plusMinutes(20)
            println(value.toString())
        }
    "#);
    assert_eq!(out, &["02:05"]);
}

#[test]
fn test_local_time_is_before() {
    let out = run_prints(r#"
        fun main() {
            val a = java.time.LocalTime.parse("08:00:00")
            val b = java.time.LocalTime.parse("08:00:01")
            println(a.isBefore(b))
            println(a == b.minusSeconds(1))
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_local_datetime_split_and_combine() {
    let out = run_prints(r#"
        fun main() {
            val value = java.time.LocalDateTime.parse("2024-07-01T10:11:12")
            println(value.date.toString())
            println(value.toLocalTime().toString())
        }
    "#);
    assert_eq!(out, &["2024-07-01", "10:11:12"]);
}

#[test]
fn test_local_datetime_plus_days_and_hours() {
    let out = run_prints(r#"
        fun main() {
            val value = java.time.LocalDateTime.parse("2024-07-01T10:00:00").plusDays(2).plusHours(5)
            println(value.toString())
        }
    "#);
    assert_eq!(out, &["2024-07-03T15:00"]);
}

#[test]
fn test_local_datetime_compare_equal() {
    let out = run_prints(r#"
        fun main() {
            val a = java.time.LocalDateTime.parse("2024-01-01T00:00")
            val b = java.time.LocalDateTime.parse("2024-01-01T00:00")
            println(a == b)
            println(a.compareTo(b))
        }
    "#);
    assert_eq!(out, &["true", "0"]);
}

#[test]
fn test_duration_between_local_time() {
    let out = run_prints(r#"
        fun main() {
            val start = java.time.LocalTime.parse("08:00")
            val end = java.time.LocalTime.parse("10:30")
            val d = java.time.Duration.between(start, end)
            println(d.toMinutes())
            println(d.seconds)
        }
    "#);
    assert_eq!(out, &["150", "9000"]);
}

#[test]
fn test_duration_parse_and_multiply() {
    let out = run_prints(r#"
        fun main() {
            val d = java.time.Duration.parse("PT2H30M")
            println(d.toHours())
            println(d.minusHours(1).toMinutes())
        }
    "#);
    assert_eq!(out, &["2", "90"]);
}

#[test]
fn test_duration_plus_and_minus() {
    let out = run_prints(r#"
        fun main() {
            val d = java.time.Duration.ofHours(1).plusMinutes(90).minusMinutes(15)
            println(d.toMinutes())
        }
    "#);
    assert_eq!(out, &["135"]);
}

#[test]
fn test_duration_negation_and_zero() {
    let out = run_prints(r#"
        fun main() {
            val d = java.time.Duration.ofMinutes(10).negated()
            println(d.toMinutes())
            println((d + java.time.Duration.ofMinutes(10)).isZero())
        }
    "#);
    assert_eq!(out, &["-10", "true"]);
}

#[test]
fn test_period_between_dates() {
    let out = run_prints(r#"
        fun main() {
            val start = java.time.LocalDate.parse("2024-01-01")
            val end = java.time.LocalDate.parse("2024-03-11")
            val p = java.time.Period.between(start, end)
            println(p.months)
            println(p.days)
            println(p.years)
        }
    "#);
    assert_eq!(out, &["2", "10", "0"]);
}

#[test]
fn test_period_days_across_months() {
    let out = run_prints(r#"
        fun main() {
            val start = java.time.LocalDate.parse("2023-11-30")
            val end = java.time.LocalDate.parse("2023-12-01")
            val p = java.time.Period.between(start, end)
            println(p.days)
            println(p.months)
            println(p.years)
        }
    "#);
    assert_eq!(out, &["1", "0", "0"]);
}

#[test]
fn test_zoned_date_time_parse_with_zone() {
    let out = run_prints(r#"
        fun main() {
            val value = java.time.ZonedDateTime.parse("2024-01-01T10:15:30+01:00[Europe/Paris]")
            println(value.zone.id)
            println(value.offset)
            println(value.toLocalDateTime().toString())
        }
    "#);
    assert_eq!(out, &["Europe/Paris", "+01:00", "2024-01-01T10:15:30"]);
}

#[test]
fn test_zoned_date_time_with_fixed_offset() {
    let out = run_prints(r#"
        fun main() {
            val zone = java.time.ZoneId.of("UTC")
            val value = java.time.ZonedDateTime.of(
                java.time.LocalDateTime.of(2024, 1, 1, 12, 0),
                zone
            )
            println(value.toOffsetDateTime().offset.id)
            println(value.toInstant().toEpochMilli())
        }
    "#);
    assert_eq!(out, &["Z", "1704110400000"]);
}

#[test]
fn test_offset_date_time_parse() {
    let out = run_prints(r#"
        fun main() {
            val value = java.time.OffsetDateTime.parse("2024-06-01T12:00:00+02:00")
            println(value.offset.id)
            println(value.toLocalDateTime().toString())
            println(value.toInstant().toEpochMilli())
        }
    "#);
    assert_eq!(out, &["+02:00", "2024-06-01T12:00", "1717236000000"]);
}

#[test]
fn test_epoch_millis_roundtrip_with_instant() {
    let out = run_prints(r#"
        fun main() {
            val instant = java.time.Instant.ofEpochMilli(1_700_000_000_000)
            println(instant.epochSecond)
            println(java.time.Instant.ofEpochSecond(instant.epochSecond).toEpochMilli() >= 1_700_000_000_000)
        }
    "#);
    assert_eq!(out, &["1700000000", "true"]);
}

#[test]
fn test_instant_duration_between_instant_points() {
    let out = run_prints(r#"
        fun main() {
            val a = java.time.Instant.parse("2024-01-01T00:00:00Z")
            val b = java.time.Instant.parse("2024-01-01T00:00:30Z")
            val d = java.time.Duration.between(a, b)
            println(d.seconds)
            println(d.toMillis())
        }
    "#);
    assert_eq!(out, &["30", "30000"]);
}

#[test]
fn test_chrono_unit_days_between_dates() {
    let out = run_prints(r#"
        fun main() {
            val a = java.time.LocalDate.parse("2024-01-01")
            val b = java.time.LocalDate.parse("2024-01-10")
            println(java.time.temporal.ChronoUnit.DAYS.between(a, b))
            println(java.time.temporal.ChronoUnit.MONTHS.between(a, b))
            println(java.time.temporal.ChronoUnit.WEEKS.between(a, b))
        }
    "#);
    assert_eq!(out, &["9", "0", "1"]);
}

#[test]
fn test_month_day_parse_and_fields() {
    let out = run_prints(r#"
        fun main() {
            val value = java.time.MonthDay.parse("--12-25")
            println(value.monthValue)
            println(value.dayOfMonth)
        }
    "#);
    assert_eq!(out, &["12", "25"]);
}

#[test]
fn test_year_month_parse_and_plus() {
    let out = run_prints(r#"
        fun main() {
            val value = java.time.YearMonth.parse("2024-05").plusMonths(8)
            println(value.year)
            println(value.monthValue)
        }
    "#);
    assert_eq!(out, &["2025", "1"]);
}

#[test]
fn test_week_fields_for_specific_date() {
    let out = run_prints(r#"
        fun main() {
            val value = java.time.LocalDate.parse("2024-01-02")
            val week = value.get(java.time.temporal.IsoFields.WEEK_OF_WEEK_BASED_YEAR)
            val weekYear = value.get(java.time.temporal.IsoFields.WEEK_BASED_YEAR)
            println(week)
            println(weekYear)
            println(value.dayOfWeek.value)
        }
    "#);
    assert_eq!(out, &["1", "2024", "2"]);
}

#[test]
fn test_temporal_with_year_adjustments() {
    let out = run_prints(r#"
        fun main() {
            val value = java.time.LocalDate.parse("2024-02-29").withYear(2025)
            println(value.toString())
            val atStart = value.withMonth(1).withDayOfMonth(1)
            println(atStart.toString())
        }
    "#);
    assert_eq!(out, &["2025-02-28", "2025-01-01"]);
}

#[test]
fn test_year_month_day_of_week_values() {
    let out = run_prints(r#"
        fun main() {
            val value = java.time.LocalDate.parse("2024-07-30")
            println(value.dayOfWeek.value)
            println(value.dayOfWeek.name)
        }
    "#);
    assert_eq!(out, &["2", "TUESDAY"]);
}

#[test]
fn test_advance_and_truncate_date_time() {
    let out = run_prints(r#"
        fun main() {
            val value = java.time.LocalDateTime.parse("2024-01-01T10:59:59")
            println(value.plusSeconds(62).toString())
            println(value.with(java.time.temporal.ChronoField.HOUR_OF_DAY, 0).toString())
        }
    "#);
    assert_eq!(out, &["2024-01-01T11:01:01", "2024-01-01T00:59:59"]);
}

#[test]
fn test_clock_system_now_monotonic_properties() {
    let out = run_prints(r#"
        fun main() {
            val instant = java.time.Instant.parse("2024-01-01T00:00:00Z")
            val clock = java.time.Clock.fixed(instant, java.time.ZoneId.of("UTC"))
            val a = java.time.Instant.now(clock)
            val b = java.time.Instant.now(clock)
            println(a.toString())
            println(a == b)
        }
    "#);
    assert_eq!(out, &["2024-01-01T00:00:00Z", "true"]);
}
