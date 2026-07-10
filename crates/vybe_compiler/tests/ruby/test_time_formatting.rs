use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_time_formatting_strftime_date, "puts Time.utc(2024, 2, 29).strftime('%Y-%m-%d')", "2024-02-29");
ruby_test!(test_time_formatting_strftime_time, "puts Time.utc(2024, 1, 1, 14, 30, 45).strftime('%H:%M:%S')", "14:30:45");
ruby_test!(test_time_formatting_strftime_12_hour, "puts Time.utc(2024, 1, 1, 14, 30, 45).strftime('%I:%M:%S %p')", "02:30:45 PM");
ruby_test!(test_time_formatting_strftime_weekday, "puts Time.utc(2024, 1, 1).strftime('%A %a')", "Monday Mon");
ruby_test!(test_time_formatting_strftime_month, "puts Time.utc(2024, 1, 1).strftime('%B %b')", "January Jan");
ruby_test!(test_time_formatting_strftime_timezone, "puts Time.utc(2024, 1, 1).strftime('%Z')", "UTC");
ruby_test!(test_time_formatting_to_s, "puts Time.utc(2024, 1, 1, 12, 0, 0).to_s", "2024-01-01 12:00:00 UTC");
ruby_test!(test_time_formatting_inspect, "puts Time.utc(2024, 1, 1, 12, 0, 0).inspect", "2024-01-01 12:00:00 UTC");
ruby_test!(test_time_formatting_asctime, "puts Time.utc(2024, 1, 1, 12, 0, 0).asctime", "Mon Jan  1 12:00:00 2024");
ruby_test!(test_time_formatting_ctime, "puts Time.utc(2024, 1, 1, 12, 0, 0).ctime", "Mon Jan  1 12:00:00 2024");
