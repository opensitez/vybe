//! time.Parse, Format, ParseDuration, Duration.String, Add/Sub, Before/After —
//! layout and duration semantics distinct from `test_time_package.rs` (RFC3339,
//! Unix epoch, Sleep/Tick/After compile smoke).

use crate::helpers::*;

go_run_cases! {
    // ParseDuration — units, composites, signs, fractions
    time_parse_duration_one_hour => (
        "package main; import \"fmt\"; import \"time\"; func main() { d, _ := time.ParseDuration(\"1h\"); fmt.Println(d.Hours()) }",
        vec!["1"]
    ),
    time_parse_duration_hours_minutes => (
        "package main; import \"fmt\"; import \"time\"; func main() { d, _ := time.ParseDuration(\"2h30m\"); fmt.Println(d.Minutes()) }",
        vec!["150"]
    ),
    time_parse_duration_milliseconds => (
        "package main; import \"fmt\"; import \"time\"; func main() { d, _ := time.ParseDuration(\"250ms\"); fmt.Println(d.Milliseconds()) }",
        vec!["250"]
    ),
    time_parse_duration_microseconds => (
        "package main; import \"fmt\"; import \"time\"; func main() { d, _ := time.ParseDuration(\"10us\"); fmt.Println(d.Microseconds()) }",
        vec!["10"]
    ),
    time_parse_duration_negative_seconds => (
        "package main; import \"fmt\"; import \"time\"; func main() { d, _ := time.ParseDuration(\"-90s\"); fmt.Println(d.Seconds()) }",
        vec!["-90"]
    ),
    time_parse_duration_fractional_seconds => (
        "package main; import \"fmt\"; import \"time\"; func main() { d, _ := time.ParseDuration(\"1.5s\"); fmt.Println(d.Seconds()) }",
        vec!["1.5"]
    ),

    // Duration.String and unit accessors — beyond Second.String in test_time_package
    time_duration_string_hour => (
        "package main; import \"fmt\"; import \"time\"; func main() { fmt.Println(time.Hour.String()) }",
        vec!["1h0m0s"]
    ),
    time_duration_string_minute => (
        "package main; import \"fmt\"; import \"time\"; func main() { fmt.Println(time.Minute.String()) }",
        vec!["1m0s"]
    ),
    time_duration_string_millisecond => (
        "package main; import \"fmt\"; import \"time\"; func main() { fmt.Println(time.Millisecond.String()) }",
        vec!["1ms"]
    ),
    time_duration_nanoseconds_per_second => (
        "package main; import \"fmt\"; import \"time\"; func main() { fmt.Println(time.Second.Nanoseconds()) }",
        vec!["1000000000"]
    ),
    time_duration_hours_from_minutes => (
        "package main; import \"fmt\"; import \"time\"; func main() { d := 90 * time.Minute; fmt.Println(d.Hours()) }",
        vec!["1.5"]
    ),
    time_duration_round_trip_string => (
        "package main; import \"fmt\"; import \"time\"; func main() { d := 3 * time.Hour; d2, _ := time.ParseDuration(d.String()); fmt.Println(d2.Hours()) }",
        vec!["3"]
    ),

    // Format — reference layouts and stdlib constants (not RFC3339)
    time_format_custom_date_layout => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2020, 3, 15, 14, 30, 0, 0, time.UTC); fmt.Println(t.Format(\"2006-01-02\")) }",
        vec!["2020-03-15"]
    ),
    time_format_custom_time_layout => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2020, 3, 15, 14, 30, 5, 0, time.UTC); fmt.Println(t.Format(\"15:04:05\")) }",
        vec!["14:30:05"]
    ),
    time_format_rfc822_utc => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2020, 3, 15, 14, 30, 0, 0, time.UTC); fmt.Println(t.Format(time.RFC822)) }",
        vec!["15 Mar 20 14:30 UTC"]
    ),
    time_format_kitchen_pm => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2020, 3, 15, 14, 30, 0, 0, time.UTC); fmt.Println(t.Format(time.Kitchen)) }",
        vec!["2:30PM"]
    ),
    time_format_unix_date_constant => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2020, 3, 15, 14, 30, 0, 0, time.UTC); fmt.Println(t.Format(time.UnixDate)) }",
        vec!["Sun Mar 15 14:30:00 UTC 2020"]
    ),
    time_format_stamp_micro_constant => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2020, 3, 15, 14, 30, 5, 123456000, time.UTC); fmt.Println(t.Format(time.StampMicro)) }",
        vec!["Mar 15 14:30:05.123456"]
    ),

    // Parse — custom layouts and stdlib constants (not RFC3339)
    time_parse_custom_date_layout => (
        "package main; import \"fmt\"; import \"time\"; func main() { t, _ := time.Parse(\"2006-01-02\", \"2019-07-04\"); fmt.Println(t.Year()) }",
        vec!["2019"]
    ),
    time_parse_custom_datetime_layout => (
        "package main; import \"fmt\"; import \"time\"; func main() { t, _ := time.Parse(\"2006-01-02 15:04:05\", \"2020-01-02 03:04:05\"); fmt.Println(t.Hour()) }",
        vec!["3"]
    ),
    time_parse_rfc822_utc => (
        "package main; import \"fmt\"; import \"time\"; func main() { t, _ := time.Parse(time.RFC822, \"02 Jan 20 03:04 UTC\"); fmt.Println(t.Day()) }",
        vec!["2"]
    ),
    time_parse_stamp_constant => (
        "package main; import \"fmt\"; import \"time\"; func main() { t, _ := time.Parse(time.Stamp, \"Mar 15 14:30:05\"); fmt.Println(t.Month()) }",
        vec!["March"]
    ),
    time_parse_unix_date_constant => (
        "package main; import \"fmt\"; import \"time\"; func main() { t, _ := time.Parse(time.UnixDate, \"Sun Mar 15 14:30:00 UTC 2020\"); fmt.Println(t.Day()) }",
        vec!["15"]
    ),

    // Add, AddDate, Sub — duration arithmetic beyond basic Unix offsets
    time_add_twenty_four_hours => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2020, 1, 1, 0, 0, 0, 0, time.UTC); later := t.Add(24 * time.Hour); fmt.Println(later.Day()) }",
        vec!["2"]
    ),
    time_add_negative_duration => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2020, 1, 2, 12, 0, 0, 0, time.UTC); earlier := t.Add(-6 * time.Hour); fmt.Println(earlier.Hour()) }",
        vec!["6"]
    ),
    time_add_date_month_rollover => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2020, 1, 31, 0, 0, 0, 0, time.UTC); later := t.AddDate(0, 1, 0); fmt.Println(int(later.Month())) }",
        vec!["3"]
    ),
    time_sub_same_instant_zero => (
        "package main; import \"fmt\"; import \"time\"; func main() { a := time.Unix(1000, 0); b := time.Unix(1000, 0); fmt.Println(a.Sub(b).Nanoseconds()) }",
        vec!["0"]
    ),
    time_sub_nanosecond_precision => (
        "package main; import \"fmt\"; import \"time\"; func main() { a := time.Unix(0, 500000000); b := time.Unix(0, 0); fmt.Println(a.Sub(b).Nanoseconds()) }",
        vec!["500000000"]
    ),

    // Before, After, Equal — ordering edge cases beyond paired Before/After
    time_before_strictly_earlier => (
        "package main; import \"fmt\"; import \"time\"; func main() { early := time.Unix(10, 0); late := time.Unix(20, 0); fmt.Println(early.Before(late)) }",
        vec!["true"]
    ),
    time_after_strictly_later => (
        "package main; import \"fmt\"; import \"time\"; func main() { early := time.Unix(10, 0); late := time.Unix(20, 0); fmt.Println(late.After(early)) }",
        vec!["true"]
    ),
    time_before_equal_instant_false => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Unix(42, 0); fmt.Println(t.Before(t)) }",
        vec!["false"]
    ),
    time_after_equal_instant_false => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Unix(42, 0); fmt.Println(t.After(t)) }",
        vec!["false"]
    ),
    time_equal_same_unix_nano => (
        "package main; import \"fmt\"; import \"time\"; func main() { a := time.Unix(5, 100); b := time.Unix(5, 100); fmt.Println(a.Equal(b)) }",
        vec!["true"]
    ),
}

go_compile_cases! {
    time_parse_in_location_compile => "package main; import \"time\"; func main() { _, _ = time.ParseInLocation(\"2006-01-02\", \"2020-01-01\", time.UTC) }",
    time_duration_round_compile => "package main; import \"time\"; func main() { _ = (3 * time.Hour).Round(time.Minute) }",
    time_in_location_compile => "package main; import \"time\"; func main() { t := time.Now(); _ = t.In(time.UTC) }",
}
