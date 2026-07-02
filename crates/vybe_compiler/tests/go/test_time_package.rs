//! time package: Duration, Unix timestamps, Format/Parse layouts.


go_run_cases! {
    time_duration_string_second => ("package main; import \"fmt\"; import \"time\"; func main() { fmt.Println(time.Second.String()) }", vec!["1s"]),
    time_duration_minutes => ("package main; import \"fmt\"; import \"time\"; func main() { d := 2 * time.Minute; fmt.Println(d.Minutes()) }", vec!["2"]),
    time_unix_zero_epoch => ("package main; import \"fmt\"; import \"time\"; func main() { t := time.Unix(0, 0).UTC(); fmt.Println(t.Year()) }", vec!["1970"]),
    time_format_rfc3339_utc => ("package main; import \"fmt\"; import \"time\"; func main() { t := time.Date(2020, 1, 2, 3, 4, 5, 0, time.UTC); fmt.Println(t.Format(time.RFC3339)) }", vec!["2020-01-02T03:04:05Z"]),
    time_parse_rfc3339 => ("package main; import \"fmt\"; import \"time\"; func main() { t, _ := time.Parse(time.RFC3339, \"2021-06-15T12:00:00Z\"); fmt.Println(t.Month()) }", vec!["June"]),
    time_add_duration => ("package main; import \"fmt\"; import \"time\"; func main() { t := time.Unix(100, 0); later := t.Add(10 * time.Second); fmt.Println(later.Unix()) }", vec!["110"]),
    time_sub_duration => ("package main; import \"fmt\"; import \"time\"; func main() { a := time.Unix(100, 0); b := time.Unix(40, 0); fmt.Println(a.Sub(b).Seconds()) }", vec!["60"]),
    time_before_after => ("package main; import \"fmt\"; import \"time\"; func main() { early := time.Unix(1,0); late := time.Unix(2,0); fmt.Println(early.Before(late)); fmt.Println(late.After(early)) }", vec!["true", "true"]),
}

go_compile_cases! {
    time_sleep_compile => "package main; import \"time\"; func main() { time.Sleep(time.Nanosecond) }",
    time_tick_compile => "package main; import \"time\"; func main() { _ = time.Tick(time.Second) }",
    time_after_compile => "package main; import \"time\"; func main() { _ = time.After(time.Second) }",
    time_location_load => "package main; import \"time\"; func main() { _, _ = time.LoadLocation(\"UTC\") }",
    time_date_components => "package main; import \"time\"; func main() { _ = time.Date(2020, 12, 25, 0, 0, 0, 0, time.UTC) }",
}
