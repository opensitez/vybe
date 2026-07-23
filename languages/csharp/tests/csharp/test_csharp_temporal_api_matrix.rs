use super::helpers::run_csharp;

fn bool_text(v: bool) -> &'static str {
    if v { "True" } else { "False" }
}

#[test]
fn datetime_matrix_add_days_boundaries() {
    let cases = [
        (2020, 1, 31, 1, 2020, 2, 1),
        (2019, 1, 31, 1, 2019, 2, 1),
        (2020, 2, 28, 1, 2020, 2, 29),
        (2019, 12, 31, 1, 2020, 1, 1),
        (2020, 3, 1, -1, 2020, 2, 29),
        (2019, 3, 1, -1, 2019, 2, 28),
    ];

    for (year, month, day, delta, expected_year, expected_month, expected_day) in cases {
        let src = format!(
            "var origin = new System.DateTime({year}, {month}, {day}); var moved = origin.AddDays({delta}); Console.WriteLine(moved.Year); Console.WriteLine(moved.Month); Console.WriteLine(moved.Day); Console.WriteLine((int)((moved - origin).TotalDays));"
        );
        assert_eq!(
            run_csharp(&src),
            vec![
                expected_year.to_string(),
                expected_month.to_string(),
                expected_day.to_string(),
                delta.to_string()
            ]
        );
    }
}

#[test]
fn datetime_matrix_add_months_boundaries() {
    let cases = [
        (2020, 1, 31, 1, 2020, 2, 29),
        (2019, 1, 31, 1, 2019, 2, 28),
        (2019, 8, 31, 6, 2020, 2, 29),
        (2021, 10, 31, 2, 2021, 12, 31),
        (2021, 12, 31, -3, 2021, 9, 30),
    ];

    for (year, month, day, delta, expected_year, expected_month, expected_day) in cases {
        let src = format!(
            "var origin = new System.DateTime({year}, {month}, {day}); var moved = origin.AddMonths({delta}); Console.WriteLine(moved.Year); Console.WriteLine(moved.Month); Console.WriteLine(moved.Day);"
        );
        assert_eq!(
            run_csharp(&src),
            vec![
                expected_year.to_string(),
                expected_month.to_string(),
                expected_day.to_string()
            ]
        );
    }
}

#[test]
fn timespan_matrix_total_units() {
    let cases = [0, 1, 30, 60, 90, 120, 3600, 86400];

    for seconds in cases {
        let src = format!(
            "var span = System.TimeSpan.FromSeconds({seconds}); Console.WriteLine((long)span.TotalSeconds); Console.WriteLine((long)span.TotalMinutes); Console.WriteLine((long)span.TotalHours);"
        );
        assert_eq!(
            run_csharp(&src),
            vec![
                seconds.to_string(),
                (seconds / 60).to_string(),
                (seconds / 3600).to_string(),
            ]
        );
    }
}

#[test]
fn timespan_matrix_addition_and_compare() {
    let cases = [(2, 30), (5, 90), (15, 45), (120, 90), (60, 15), (300, 600)];

    for (minutes, add_seconds) in cases {
        let src = format!(
            "var left = System.TimeSpan.FromMinutes({minutes}); var right = System.TimeSpan.FromSeconds({add_seconds}); var sum = left + right; Console.WriteLine((long)sum.TotalSeconds); Console.WriteLine((left + right == left) ? 1 : 0);"
        );
        let expected_sum = (minutes * 60 + add_seconds).to_string();
        assert_eq!(
            run_csharp(&src),
            vec![expected_sum, bool_text(false).to_string()]
        );
    }
}

#[test]
fn timespan_matrix_comparison_matrix() {
    let cases = [(30, 30), (60, 90), (180, 90), (45, 45)];

    for (left_sec, right_sec) in cases {
        let src = format!(
            "var left = System.TimeSpan.FromSeconds({left_sec}); var right = System.TimeSpan.FromSeconds({right_sec}); int cmp = left.CompareTo(right); bool eq = left == right; Console.WriteLine(cmp > 0 ? 1 : (cmp < 0 ? -1 : 0)); Console.WriteLine(eq);"
        );
        let expected_cmp = match left_sec.cmp(&right_sec) {
            std::cmp::Ordering::Less => -1,
            std::cmp::Ordering::Equal => 0,
            std::cmp::Ordering::Greater => 1,
        };
        assert_eq!(
            run_csharp(&src),
            vec![
                expected_cmp.to_string(),
                bool_text(left_sec == right_sec).to_string()
            ]
        );
    }
}

#[test]
fn datetime_matrix_time_components() {
    let cases = [(2020, 2, 3, 4, 5, 6), (2019, 12, 31, 23, 59, 58)];

    for (year, month, day, hour, minute, second) in cases {
        let src = format!(
            "var dt = new System.DateTime({year}, {month}, {day}, {hour}, {minute}, {second}); Console.WriteLine(dt.Year); Console.WriteLine(dt.Month); Console.WriteLine(dt.Day); Console.WriteLine(dt.Hour); Console.WriteLine(dt.Minute); Console.WriteLine(dt.Second);"
        );
        assert_eq!(
            run_csharp(&src),
            vec![
                year.to_string(),
                month.to_string(),
                day.to_string(),
                hour.to_string(),
                minute.to_string(),
                second.to_string()
            ]
        );
    }
}
