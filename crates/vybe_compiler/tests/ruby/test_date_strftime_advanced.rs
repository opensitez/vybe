macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_date_strftime_iso8601,
    "require 'date'; puts Date.new(2001, 2, 3).strftime('%F')",
    "2001-02-03"
);
ruby_test!(
    test_date_strftime_year_month_day,
    "require 'date'; puts Date.new(2001, 2, 3).strftime('%Y-%m-%d')",
    "2001-02-03"
);
ruby_test!(
    test_date_strftime_short_year,
    "require 'date'; puts Date.new(2001, 2, 3).strftime('%y')",
    "01"
);
ruby_test!(
    test_date_strftime_weekday_name,
    "require 'date'; puts Date.new(2001, 2, 3).strftime('%A')",
    "Saturday"
);
ruby_test!(
    test_date_strftime_short_weekday,
    "require 'date'; puts Date.new(2001, 2, 3).strftime('%a')",
    "Sat"
);
ruby_test!(
    test_date_strftime_month_name,
    "require 'date'; puts Date.new(2001, 2, 3).strftime('%B')",
    "February"
);
ruby_test!(
    test_date_strftime_short_month,
    "require 'date'; puts Date.new(2001, 2, 3).strftime('%b')",
    "Feb"
);
ruby_test!(
    test_date_strftime_day_of_year,
    "require 'date'; puts Date.new(2001, 2, 3).strftime('%j')",
    "034"
);
ruby_test!(
    test_time_strftime_time_components,
    "puts Time.utc(2001, 2, 3, 4, 5, 6).strftime('%H:%M:%S')",
    "04:05:06"
);
ruby_test!(
    test_time_strftime_am_pm,
    "puts Time.utc(2001, 2, 3, 16, 5, 6).strftime('%I %p')",
    "04 PM"
);
ruby_test!(
    test_time_strftime_timezone,
    "puts Time.utc(2001, 2, 3).strftime('%Z')",
    "UTC"
);
