use crate::helpers::run_prints;

#[test]
fn test_kotlin_time_zero_duration_is_zero() {
    let out = run_prints(
        r#"
        fun main() {
            println(Duration.ZERO.toDouble(DurationUnit.SECONDS))
            println(Duration.ZERO.inWholeMilliseconds)
        }
    "#,
    );
    assert_eq!(out, &["0.0", "0"]);
}

#[test]
fn test_kotlin_time_seconds_unit_conversion() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 5.toDuration(DurationUnit.SECONDS)
            println(value.inWholeMilliseconds)
            println(value.inWholeSeconds)
        }
    "#,
    );
    assert_eq!(out, &["5000", "5"]);
}

#[test]
fn test_kotlin_time_milliseconds_unit_conversion() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 1500.toDuration(DurationUnit.MILLISECONDS)
            println(value.inWholeSeconds)
            println(value.inWholeMilliseconds)
        }
    "#,
    );
    assert_eq!(out, &["1", "1500"]);
}

#[test]
fn test_kotlin_time_minutes_unit_conversion() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 2.toDuration(DurationUnit.MINUTES)
            println(value.inWholeSeconds)
            println(value.inWholeMilliseconds)
            println(value.inWholeMinutes)
        }
    "#,
    );
    assert_eq!(out, &["120", "120000", "2"]);
}

#[test]
fn test_kotlin_time_hours_unit_conversion() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 1.toDuration(DurationUnit.HOURS)
            println(value.inWholeMinutes)
            println(value.inWholeSeconds)
        }
    "#,
    );
    assert_eq!(out, &["60", "3600"]);
}

#[test]
fn test_kotlin_time_days_unit_conversion() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 1.toDuration(DurationUnit.DAYS)
            println(value.inWholeHours)
            println(value.inWholeMinutes)
            println(value.inWholeSeconds)
        }
    "#,
    );
    assert_eq!(out, &["24", "1440", "86400"]);
}

#[test]
fn test_kotlin_time_duration_addition() {
    let out = run_prints(
        r#"
        fun main() {
            val left = 2.toDuration(DurationUnit.SECONDS)
            val right = 500.toDuration(DurationUnit.MILLISECONDS)
            println((left + right).inWholeMilliseconds)
        }
    "#,
    );
    assert_eq!(out, &["2500"]);
}

#[test]
fn test_kotlin_time_duration_subtraction() {
    let out = run_prints(
        r#"
        fun main() {
            val left = 5.toDuration(DurationUnit.SECONDS)
            val right = 2.toDuration(DurationUnit.SECONDS)
            println((left - right).inWholeSeconds)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_kotlin_time_duration_multiplication() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 3.toDuration(DurationUnit.SECONDS)
            println((value * 2).inWholeMilliseconds)
            println((value * 3).inWholeMilliseconds)
        }
    "#,
    );
    assert_eq!(out, &["6000", "9000"]);
}

#[test]
fn test_kotlin_time_duration_division_int() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 10.toDuration(DurationUnit.SECONDS)
            println((value / 2).inWholeSeconds)
            println((value / 5).inWholeMilliseconds)
        }
    "#,
    );
    assert_eq!(out, &["5", "2000"]);
}

#[test]
fn test_kotlin_time_duration_negation_sign() {
    let out = run_prints(
        r#"
        fun main() {
            val value = -(2.toDuration(DurationUnit.SECONDS))
            println(value.inWholeSeconds)
            println(value.isNegative())
            println(value < Duration.ZERO)
        }
    "#,
    );
    assert_eq!(out, &["-2", "true", "true"]);
}

#[test]
fn test_kotlin_time_duration_abs_via_zero_minus() {
    let out = run_prints(
        r#"
        fun main() {
            val value = -(4.toDuration(DurationUnit.SECONDS))
            val inverted = -value
            println(inverted.inWholeMilliseconds)
            println(inverted == 4.toDuration(DurationUnit.SECONDS))
        }
    "#,
    );
    assert_eq!(out, &["4000", "true"]);
}

#[test]
fn test_kotlin_time_duration_compare_equal() {
    let out = run_prints(
        r#"
        fun main() {
            val a = 1.toDuration(DurationUnit.MINUTES)
            val b = 60.toDuration(DurationUnit.SECONDS)
            println(a == b)
            println(a != b)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_kotlin_time_duration_compare_greater() {
    let out = run_prints(
        r#"
        fun main() {
            val a = 90.toDuration(DurationUnit.SECONDS)
            val b = 1.toDuration(DurationUnit.MINUTES)
            println(a > b)
            println(a >= b)
            println(a < b)
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_kotlin_time_duration_compare_with_zero() {
    let out = run_prints(
        r#"
        fun main() {
            val positive = 1.toDuration(DurationUnit.SECONDS)
            val zero = Duration.ZERO
            val negative = -(500.toDuration(DurationUnit.MILLISECONDS))
            println(positive > zero)
            println(zero > negative)
            println(negative < zero)
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_kotlin_time_duration_to_double_seconds() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 1500.toDuration(DurationUnit.MILLISECONDS)
            println(value.toDouble(DurationUnit.MILLISECONDS))
            println(value.toDouble(DurationUnit.SECONDS))
        }
    "#,
    );
    assert_eq!(out, &["1500", "1.5"]);
}

#[test]
fn test_kotlin_time_duration_to_long_nanoseconds() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 250.toDuration(DurationUnit.MILLISECONDS)
            println(value.toLong(DurationUnit.NANOSECONDS))
            println(value.toLong(DurationUnit.MICROSECONDS))
        }
    "#,
    );
    assert_eq!(out, &["250000000", "250000"]);
}

#[test]
fn test_kotlin_time_duration_truncation_floor() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 1501.toDuration(DurationUnit.MILLISECONDS)
            println(value.inWholeSeconds)
            println(value.inWholeMilliseconds)
        }
    "#,
    );
    assert_eq!(out, &["1", "1501"]);
}

#[test]
fn test_kotlin_time_duration_is_finite_for_real() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 10.toDuration(DurationUnit.SECONDS)
            println(value.isInfinite())
            println(value.isFinite())
        }
    "#,
    );
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_kotlin_time_duration_is_infinite() {
    let out = run_prints(
        r#"
        fun main() {
            val value = Duration.INFINITE
            println(value.isInfinite())
            println(value.isFinite())
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_kotlin_time_infinite_addition_still_infinite() {
    let out = run_prints(
        r#"
        fun main() {
            val value = Duration.INFINITE + 10.toDuration(DurationUnit.SECONDS)
            println(value.isInfinite())
            println(Duration.INFINITE - 10.toDuration(DurationUnit.SECONDS) == Duration.INFINITE)
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_kotlin_time_zero_plus_zero_stable() {
    let out = run_prints(
        r#"
        fun main() {
            val value = Duration.ZERO + Duration.ZERO
            println(value == Duration.ZERO)
            println(value.inWholeMilliseconds)
        }
    "#,
    );
    assert_eq!(out, &["true", "0"]);
}

#[test]
fn test_kotlin_time_subtract_to_negative() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 1.toDuration(DurationUnit.SECONDS) - 2.toDuration(DurationUnit.SECONDS)
            println(value.inWholeSeconds)
            println(value < Duration.ZERO)
        }
    "#,
    );
    assert_eq!(out, &["-1", "true"]);
}

#[test]
fn test_kotlin_time_negative_plus_positive_cancel_to_zero() {
    let out = run_prints(
        r#"
        fun main() {
            val value = -(3.toDuration(DurationUnit.SECONDS)) + 3.toDuration(DurationUnit.SECONDS)
            println(value == Duration.ZERO)
            println(value.inWholeMilliseconds)
        }
    "#,
    );
    assert_eq!(out, &["true", "0"]);
}

#[test]
fn test_kotlin_time_non_trivial_unit_rounding_seconds_to_ms() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 2.toDuration(DurationUnit.SECONDS)
            println(value.toLong(DurationUnit.MILLISECONDS))
            println(value.toLong(DurationUnit.MICROSECONDS))
        }
    "#,
    );
    assert_eq!(out, &["2000", "2000000"]);
}

#[test]
fn test_kotlin_time_non_trivial_unit_rounding_ms_to_s() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 1999.toDuration(DurationUnit.MILLISECONDS)
            println(value.toLong(DurationUnit.SECONDS))
            println(value.toDouble(DurationUnit.SECONDS))
        }
    "#,
    );
    assert_eq!(out, &["1", "1.999"]);
}

#[test]
fn test_kotlin_time_duration_range_check_small() {
    let out = run_prints(
        r#"
        fun main() {
            val a = 500.toDuration(DurationUnit.MILLISECONDS)
            val b = 1.toDuration(DurationUnit.SECONDS)
            val c = 1500.toDuration(DurationUnit.MILLISECONDS)
            println(a < b)
            println(c > b)
            println(c > a)
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_kotlin_time_duration_between_chained_ops() {
    let out = run_prints(
        r#"
        fun main() {
            val value = (1.toDuration(DurationUnit.MINUTES) + 30.toDuration(DurationUnit.SECONDS)) - 1.toDuration(DurationUnit.SECONDS)
            println(value.inWholeSeconds)
        }
    "#,
    );
    assert_eq!(out, &["89"]);
}

#[test]
fn test_kotlin_time_duration_mixed_sign_chain() {
    let out = run_prints(
        r#"
        fun main() {
            val value = (5.toDuration(DurationUnit.SECONDS) - 2.toDuration(DurationUnit.SECONDS)) + -(1.toDuration(DurationUnit.SECONDS))
            println(value.inWholeSeconds)
            println(value == 2.toDuration(DurationUnit.SECONDS))
        }
    "#,
    );
    assert_eq!(out, &["2", "true"]);
}

#[test]
fn test_kotlin_time_duration_scale_round_trip() {
    let out = run_prints(
        r#"
        fun main() {
            val base = 12.toDuration(DurationUnit.MILLISECONDS)
            println((base * 5).toLong(DurationUnit.MILLISECONDS))
            println(((base * 5) / 5).inWholeMilliseconds)
            println(base.inWholeMilliseconds)
        }
    "#,
    );
    assert_eq!(out, &["60", "12", "12"]);
}

#[test]
fn test_kotlin_time_duration_fraction_to_long_floor_behavior() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 2500.toDuration(DurationUnit.MILLISECONDS)
            println(value.toLong(DurationUnit.SECONDS))
            println(value.toDouble(DurationUnit.SECONDS))
        }
    "#,
    );
    assert_eq!(out, &["2", "2.5"]);
}

#[test]
fn test_kotlin_time_duration_components_like_whole_minutes() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 3700.toDuration(DurationUnit.SECONDS)
            println(value.inWholeMinutes)
            println(value.inWholeHours)
            println(value.inWholeDays)
        }
    "#,
    );
    assert_eq!(out, &["61", "1", "0"]);
}

#[test]
fn test_kotlin_time_duration_to_string_is_string_like() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 90.toDuration(DurationUnit.SECONDS)
            println(value.toString().contains("s"))
            println(value.inWholeMinutes)
        }
    "#,
    );
    assert_eq!(out, &["true", "1"]);
}
