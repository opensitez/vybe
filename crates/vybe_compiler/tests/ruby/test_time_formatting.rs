macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_time_strftime_basic,
    "t = Time.utc(2023, 1, 2, 15, 4, 5); puts t.strftime('%Y-%m-%d %H:%M:%S')",
    "2023-01-02 15:04:05"
);
ruby_test!(
    test_time_strftime_zone,
    "t = Time.utc(2023, 1, 2); puts t.strftime('%Z')",
    "UTC"
);
ruby_test!(
    test_time_strftime_weekday,
    "t = Time.utc(2023, 1, 2); puts t.strftime('%A')",
    "Monday"
);
ruby_test!(
    test_time_strftime_month,
    "t = Time.utc(2023, 1, 2); puts t.strftime('%B')",
    "January"
);
ruby_test!(
    test_time_iso8601,
    "require 'time'; t = Time.utc(2023, 1, 2, 15, 4, 5); puts t.iso8601",
    "2023-01-02T15:04:05Z"
);
ruby_test!(
    test_time_rfc2822,
    "require 'time'; t = Time.utc(2023, 1, 2, 15, 4, 5); puts t.rfc2822",
    "Mon, 02 Jan 2023 15:04:05 -0000"
);
ruby_test!(
    test_time_httpdate,
    "require 'time'; t = Time.utc(2023, 1, 2, 15, 4, 5); puts t.httpdate",
    "Mon, 02 Jan 2023 15:04:05 GMT"
);
