//! time.LoadLocation, FixedZone, Date, UnixMilli/Micro/Nano, Truncate/Round,
//! Weekday/Month constants, Since/Until — distinct from `test_time_package.rs`
//! (RFC3339, basic Unix, Sleep/Tick) and `test_time_parse_format.rs` (ParseDuration,
//! custom Format/Parse layouts, Add/Sub duration).


go_run_cases! {
    time_date_utc_components => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2024, time.March, 15, 14, 30, 45, 0, time.UTC); fmt.Println(t.Year()); fmt.Println(t.Month()); fmt.Println(t.Day()); fmt.Println(t.Hour()) }",
        vec!["2024", "March", "15", "14"]
    ),
    time_date_zero_nano => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2000, 1, 1, 0, 0, 0, 0, time.UTC); fmt.Println(t.Nanosecond()); fmt.Println(t.Second()) }",
        vec!["0", "0"]
    ),
    time_date_with_nanoseconds => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2020, 6, 1, 0, 0, 0, 123456789, time.UTC); fmt.Println(t.Nanosecond()) }",
        vec!["123456789"]
    ),
    time_date_leap_year_feb => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2024, time.February, 29, 12, 0, 0, 0, time.UTC); fmt.Println(t.Month()); fmt.Println(t.Day()) }",
        vec!["February", "29"]
    ),
    time_date_year_boundary => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(1999, time.December, 31, 23, 59, 59, 0, time.UTC); fmt.Println(t.Year()); fmt.Println(t.Month()) }",
        vec!["1999", "December"]
    ),

    time_unix_milli_epoch => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.UnixMilli(0); fmt.Println(t.Unix()) }",
        vec!["0"]
    ),
    time_unix_milli_seconds => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.UnixMilli(1500); fmt.Println(t.Unix()); fmt.Println(t.UnixMilli()) }",
        vec!["1", "1500"]
    ),
    time_unix_micro_epoch => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.UnixMicro(0); fmt.Println(t.UnixMicro()) }",
        vec!["0"]
    ),
    time_unix_micro_value => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.UnixMicro(1234567); fmt.Println(t.Unix()); fmt.Println(t.UnixMicro()) }",
        vec!["1", "1234567"]
    ),
    time_unix_nano_epoch => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Unix(1, 500); fmt.Println(t.UnixNano()) }",
        vec!["1000000500"]
    ),
    time_unix_nano_from_unix_nano => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Unix(0, 999999999); fmt.Println(t.Nanosecond()) }",
        vec!["999999999"]
    ),
    time_unix_milli_negative => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.UnixMilli(-1000); fmt.Println(t.Unix()) }",
        vec!["-1"]
    ),

    time_fixed_zone_positive_offset => (
        "package main; import \"fmt\"; import \"time\"; func main() { loc := time.FixedZone(\"EST\", -5*3600); t := time.Date(2020, 1, 1, 12, 0, 0, 0, loc); fmt.Println(t.Location().String()); fmt.Println(t.Hour()) }",
        vec!["EST", "12"]
    ),
    time_fixed_zone_zero_offset => (
        "package main; import \"fmt\"; import \"time\"; func main() { loc := time.FixedZone(\"UTC+0\", 0); t := time.Date(2020, 1, 1, 0, 0, 0, 0, loc); fmt.Println(t.UTC().Hour()) }",
        vec!["0"]
    ),
    time_fixed_zone_name => (
        "package main; import \"fmt\"; import \"time\"; func main() { loc := time.FixedZone(\"Custom\", 3600); fmt.Println(loc.String()) }",
        vec!["Custom"]
    ),
    time_fixed_zone_half_hour => (
        "package main; import \"fmt\"; import \"time\"; func main() { loc := time.FixedZone(\"IST\", 5*3600+30*60); t := time.Date(2021, 7, 1, 10, 0, 0, 0, loc); fmt.Println(t.Hour()) }",
        vec!["10"]
    ),

    time_load_location_utc => (
        "package main; import \"fmt\"; import \"time\"; func main() { loc, err := time.LoadLocation(\"UTC\"); fmt.Println(err == nil); fmt.Println(loc.String()) }",
        vec!["true", "UTC"]
    ),
    time_load_location_local => (
        "package main; import \"fmt\"; import \"time\"; func main() { loc, err := time.LoadLocation(\"Local\"); fmt.Println(err == nil); fmt.Println(loc != nil) }",
        vec!["true", "true"]
    ),
    time_date_in_fixed_zone => (
        "package main; import \"fmt\"; import \"time\"; func main() { loc := time.FixedZone(\"PST\", -8*3600); t := time.Date(2022, 5, 10, 8, 0, 0, 0, loc); utc := t.UTC(); fmt.Println(utc.Hour()) }",
        vec!["16"]
    ),

    time_truncate_to_hour => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2023, 4, 5, 14, 35, 22, 0, time.UTC); truncated := t.Truncate(time.Hour); fmt.Println(truncated.Minute()); fmt.Println(truncated.Second()) }",
        vec!["0", "0"]
    ),
    time_truncate_to_day => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2023, 4, 5, 14, 35, 22, 0, time.UTC); truncated := t.Truncate(24 * time.Hour); fmt.Println(truncated.Hour()); fmt.Println(truncated.Day()) }",
        vec!["0", "5"]
    ),
    time_truncate_to_minute => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2023, 1, 1, 10, 15, 45, 0, time.UTC); truncated := t.Truncate(time.Minute); fmt.Println(truncated.Second()) }",
        vec!["0"]
    ),
    time_round_to_hour => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2023, 4, 5, 14, 35, 22, 0, time.UTC); rounded := t.Round(time.Hour); fmt.Println(rounded.Hour()) }",
        vec!["15"]
    ),
    time_round_to_hour_down => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2023, 4, 5, 14, 25, 22, 0, time.UTC); rounded := t.Round(time.Hour); fmt.Println(rounded.Hour()) }",
        vec!["14"]
    ),
    time_round_to_minute => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2023, 1, 1, 10, 15, 45, 0, time.UTC); rounded := t.Round(time.Minute); fmt.Println(rounded.Minute()); fmt.Println(rounded.Second()) }",
        vec!["16", "0"]
    ),
    time_truncate_already_aligned => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2023, 1, 1, 10, 0, 0, 0, time.UTC); truncated := t.Truncate(time.Hour); fmt.Println(truncated.Equal(t)) }",
        vec!["true"]
    ),

    time_weekday_sunday => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2023, 1, 1, 0, 0, 0, 0, time.UTC); fmt.Println(t.Weekday()) }",
        vec!["Sunday"]
    ),
    time_weekday_monday => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2023, 1, 2, 0, 0, 0, 0, time.UTC); fmt.Println(t.Weekday()) }",
        vec!["Monday"]
    ),
    time_weekday_string => (
        "package main; import \"fmt\"; import \"time\"; func main() { fmt.Println(time.Wednesday.String()) }",
        vec!["Wednesday"]
    ),
    time_weekday_constant_value => (
        "package main; import \"fmt\"; import \"time\"; func main() { fmt.Println(int(time.Friday)) }",
        vec!["5"]
    ),
    time_weekday_saturday => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2023, 1, 7, 0, 0, 0, 0, time.UTC); fmt.Println(t.Weekday()) }",
        vec!["Saturday"]
    ),

    time_month_january_constant => (
        "package main; import \"fmt\"; import \"time\"; func main() { fmt.Println(time.January) }",
        vec!["January"]
    ),
    time_month_december_constant => (
        "package main; import \"fmt\"; import \"time\"; func main() { fmt.Println(time.December) }",
        vec!["December"]
    ),
    time_month_numeric_value => (
        "package main; import \"fmt\"; import \"time\"; func main() { fmt.Println(int(time.March)) }",
        vec!["3"]
    ),
    time_month_string => (
        "package main; import \"fmt\"; import \"time\"; func main() { fmt.Println(time.April.String()) }",
        vec!["April"]
    ),
    time_month_from_date => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2020, time.July, 4, 0, 0, 0, 0, time.UTC); fmt.Println(t.Month()) }",
        vec!["July"]
    ),

    time_unix_zero_utc_year => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Unix(0, 0).UTC(); fmt.Println(t.Year()); fmt.Println(t.Month()) }",
        vec!["1970", "January"]
    ),
    time_unix_returns_seconds => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Unix(1700000000, 0); fmt.Println(t.Unix()) }",
        vec!["1700000000"]
    ),
    time_unix_milli_roundtrip => (
        "package main; import \"fmt\"; import \"time\"; func main() { ms := int64(1609459200123); t := time.UnixMilli(ms); fmt.Println(t.UnixMilli()) }",
        vec!["1609459200123"]
    ),
    time_unix_micro_roundtrip => (
        "package main; import \"fmt\"; import \"time\"; func main() { us := int64(1609459200456789); t := time.UnixMicro(us); fmt.Println(t.UnixMicro()) }",
        vec!["1609459200456789"]
    ),

    time_year_day => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2023, 1, 15, 0, 0, 0, 0, time.UTC); fmt.Println(t.YearDay()) }",
        vec!["15"]
    ),
    time_year_day_december => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2023, 12, 31, 0, 0, 0, 0, time.UTC); fmt.Println(t.YearDay()) }",
        vec!["365"]
    ),
    time_zone_offset_utc => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2020, 1, 1, 0, 0, 0, 0, time.UTC); _, off := t.Zone(); fmt.Println(off) }",
        vec!["0"]
    ),
    time_zone_offset_fixed => (
        "package main; import \"fmt\"; import \"time\"; func main() { loc := time.FixedZone(\"X\", 7200); t := time.Date(2020, 1, 1, 0, 0, 0, 0, loc); _, off := t.Zone(); fmt.Println(off) }",
        vec!["7200"]
    ),
    time_is_zero_true => (
        "package main; import \"fmt\"; import \"time\"; func main() { var t time.Time; fmt.Println(t.IsZero()) }",
        vec!["true"]
    ),
    time_is_zero_false => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Unix(1, 0); fmt.Println(t.IsZero()) }",
        vec!["false"]
    ),
    time_equal_same_instant => (
        "package main; import \"fmt\"; import \"time\"; func main() { a := time.Unix(100, 0); b := time.Unix(100, 0); fmt.Println(a.Equal(b)) }",
        vec!["true"]
    ),
    time_before_after_same_zone => (
        "package main; import \"fmt\"; import \"time\"; func main() { early := time.Date(2020, 1, 1, 0, 0, 0, 0, time.UTC); late := time.Date(2020, 1, 2, 0, 0, 0, 0, time.UTC); fmt.Println(early.Before(late)); fmt.Println(late.After(early)) }",
        vec!["true", "true"]
    ),
    time_location_utc_singleton => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Now().UTC(); fmt.Println(t.Location() == time.UTC) }",
        vec!["true"]
    ),
    time_date_month_out_of_range_normalized => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2020, 13, 1, 0, 0, 0, 0, time.UTC); fmt.Println(t.Month()); fmt.Println(t.Year()) }",
        vec!["January", "2021"]
    ),
    time_hour_minute_second => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2020, 6, 15, 9, 8, 7, 0, time.UTC); fmt.Println(t.Hour()); fmt.Println(t.Minute()); fmt.Println(t.Second()) }",
        vec!["9", "8", "7"]
    ),
    time_weekday_from_unix => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Unix(86400, 0).UTC(); fmt.Println(t.Weekday()) }",
        vec!["Friday"]
    ),
    time_month_august_constant => (
        "package main; import \"fmt\"; import \"time\"; func main() { fmt.Println(int(time.August)) }",
        vec!["8"]
    ),
    time_month_november_constant => (
        "package main; import \"fmt\"; import \"time\"; func main() { fmt.Println(time.November.String()) }",
        vec!["November"]
    ),
    time_truncate_sub_hour => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2023, 6, 1, 10, 30, 0, 0, time.UTC); truncated := t.Truncate(30 * time.Minute); fmt.Println(truncated.Minute()) }",
        vec!["30"]
    ),
    time_round_sub_hour => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2023, 6, 1, 10, 44, 0, 0, time.UTC); rounded := t.Round(30 * time.Minute); fmt.Println(rounded.Minute()) }",
        vec!["30"]
    ),
    time_fixed_zone_negative_name => (
        "package main; import \"fmt\"; import \"time\"; func main() { loc := time.FixedZone(\"MST\", -7*3600); fmt.Println(loc.String()) }",
        vec!["MST"]
    ),
    time_unix_nano_large => (
        "package main; import \"fmt\"; import \"time\"; func main() { t := time.Unix(2, 0); fmt.Println(t.UnixNano() > 0) }",
        vec!["true"]
    ),
}

go_compile_cases! {
    time_since_now => "package main; import \"time\"; func main() { _ = time.Since(time.Now()) }",
    time_until_future => "package main; import \"time\"; func main() { _ = time.Until(time.Now().Add(time.Hour)) }",
    time_since_unix_epoch => "package main; import \"time\"; func main() { _ = time.Since(time.Unix(0, 0)) }",
    time_until_fixed_date => "package main; import \"time\"; func main() { _ = time.Until(time.Date(2030, 1, 1, 0, 0, 0, 0, time.UTC)) }",
    time_load_location_america => "package main; import \"time\"; func main() { _, _ = time.LoadLocation(\"America/New_York\") }",
    time_load_location_europe => "package main; import \"time\"; func main() { _, _ = time.LoadLocation(\"Europe/London\") }",
    time_load_location_asia => "package main; import \"time\"; func main() { _, _ = time.LoadLocation(\"Asia/Tokyo\") }",
    time_date_all_months => "package main; import \"time\"; func main() { _ = time.Date(2020, time.February, 1, 0, 0, 0, 0, time.UTC); _ = time.Date(2020, time.September, 1, 0, 0, 0, 0, time.UTC) }",
    time_truncate_nanosecond => "package main; import \"time\"; func main() { t := time.Now(); _ = t.Truncate(time.Nanosecond) }",
    time_round_nanosecond => "package main; import \"time\"; func main() { t := time.Now(); _ = t.Round(time.Nanosecond) }",
    time_fixed_zone_max_offset => "package main; import \"time\"; func main() { _ = time.FixedZone(\"Max\", 14*3600) }",
    time_unix_milli_now => "package main; import \"time\"; func main() { _ = time.UnixMilli(time.Now().UnixMilli()) }",
    time_unix_micro_now => "package main; import \"time\"; func main() { _ = time.UnixMicro(time.Now().UnixMicro()) }",
    time_weekday_tuesday => "package main; import \"time\"; func main() { _ = time.Tuesday }",
    time_month_october => "package main; import \"time\"; func main() { _ = time.October }",
}
