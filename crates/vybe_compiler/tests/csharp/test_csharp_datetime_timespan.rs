use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(
    datetime_constructor_exposes_year_month_and_day,
    r#"var date = new System.DateTime(2024, 5, 17); Console.WriteLine(date.Year); Console.WriteLine(date.Month); Console.WriteLine(date.Day);"#,
    ["2024", "5", "17"]
);
csharp_case!(
    datetime_add_days_moves_to_following_date,
    r#"var date = new System.DateTime(2024, 1, 30).AddDays(2); Console.WriteLine(date.Day);"#,
    ["1"]
);
csharp_case!(
    datetime_add_hours_rolls_clock_forward,
    r#"var date = new System.DateTime(2024, 1, 1, 10, 30, 0).AddHours(5); Console.WriteLine(date.Hour);"#,
    ["15"]
);
csharp_case!(
    datetime_subtract_returns_timespan_days_delta,
    r#"var start = new System.DateTime(2024, 1, 1); var end = new System.DateTime(2024, 1, 4); Console.WriteLine((end - start).Days);"#,
    ["3"]
);
csharp_case!(
    datetime_date_property_removes_time_component,
    r#"var date = new System.DateTime(2024, 7, 8, 13, 14, 15).Date; Console.WriteLine(date.Hour);"#,
    ["0"]
);
csharp_case!(
    datetime_day_of_week_reports_expected_enum_name,
    r#"var date = new System.DateTime(2024, 6, 3); Console.WriteLine(date.DayOfWeek);"#,
    ["Monday"]
);
csharp_case!(
    datetime_days_in_month_handles_leap_year_february,
    r#"Console.WriteLine(System.DateTime.DaysInMonth(2024, 2));"#,
    ["29"]
);
csharp_case!(
    datetime_is_leap_year_recognizes_2024,
    r#"Console.WriteLine(System.DateTime.IsLeapYear(2024));"#,
    ["True"]
);
csharp_case!(
    datetime_compare_orders_earlier_before_later,
    r#"var left = new System.DateTime(2024, 1, 1); var right = new System.DateTime(2024, 1, 2); Console.WriteLine(System.DateTime.Compare(left, right));"#,
    ["-1"]
);
csharp_case!(
    datetime_add_months_crosses_year_boundary,
    r#"var date = new System.DateTime(2023, 11, 15).AddMonths(3); Console.WriteLine(date.Year); Console.WriteLine(date.Month);"#,
    ["2024", "2"]
);
csharp_case!(
    timespan_from_minutes_sets_total_seconds,
    r#"var span = System.TimeSpan.FromMinutes(2.5); Console.WriteLine(span.TotalSeconds);"#,
    ["150"]
);
csharp_case!(
    timespan_addition_combines_hours_and_minutes,
    r#"var left = System.TimeSpan.FromHours(1); var right = System.TimeSpan.FromMinutes(30); Console.WriteLine((left + right).TotalMinutes);"#,
    ["90"]
);
csharp_case!(
    timespan_subtraction_returns_remaining_minutes,
    r#"var left = System.TimeSpan.FromMinutes(45); var right = System.TimeSpan.FromMinutes(5); Console.WriteLine((left - right).TotalMinutes);"#,
    ["40"]
);
csharp_case!(
    timespan_negate_flips_sign_of_duration,
    r#"var span = System.TimeSpan.FromSeconds(9).Negate(); Console.WriteLine(span.TotalSeconds);"#,
    ["-9"]
);
csharp_case!(
    timespan_duration_returns_absolute_value,
    r#"var span = System.TimeSpan.FromSeconds(-12).Duration(); Console.WriteLine(span.TotalSeconds);"#,
    ["12"]
);
csharp_case!(
    timespan_compare_orders_shorter_before_longer,
    r#"var left = System.TimeSpan.FromSeconds(3); var right = System.TimeSpan.FromSeconds(8); Console.WriteLine(System.TimeSpan.Compare(left, right));"#,
    ["-1"]
);
csharp_case!(
    timespan_constructor_exposes_hours_minutes_and_seconds,
    r#"var span = new System.TimeSpan(2, 3, 4); Console.WriteLine(span.Hours); Console.WriteLine(span.Minutes); Console.WriteLine(span.Seconds);"#,
    ["2", "3", "4"]
);
csharp_case!(
    datetime_to_short_date_string_is_stable_for_components,
    r#"var date = new System.DateTime(2024, 12, 25); var text = date.ToShortDateString(); Console.WriteLine(text.Contains("2024"));"#,
    ["True"]
);
csharp_case!(
    datetime_time_of_day_returns_timespan_component,
    r#"var date = new System.DateTime(2024, 1, 1, 6, 45, 0); Console.WriteLine(date.TimeOfDay.TotalMinutes);"#,
    ["405"]
);
csharp_case!(
    datetime_add_timespan_combines_date_and_duration,
    r#"var date = new System.DateTime(2024, 1, 1, 1, 0, 0); var span = System.TimeSpan.FromMinutes(90); Console.WriteLine((date + span).Hour); Console.WriteLine((date + span).Minute);"#,
    ["2", "30"]
);
