macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_time_parse_iso8601,
    "require 'time'; t = Time.iso8601('2023-01-02T15:04:05Z'); puts t.utc.to_a[0..5].join('-')",
    "5-4-15-2-1-2023"
);
ruby_test!(
    test_time_parse_rfc2822,
    "require 'time'; t = Time.rfc2822('Mon, 02 Jan 2023 15:04:05 -0000'); puts t.utc.to_a[0..5].join('-')",
    "5-4-15-2-1-2023"
);
ruby_test!(
    test_time_parse_httpdate,
    "require 'time'; t = Time.httpdate('Mon, 02 Jan 2023 15:04:05 GMT'); puts t.utc.to_a[0..5].join('-')",
    "5-4-15-2-1-2023"
);
ruby_test!(
    test_time_parse_basic,
    "require 'time'; t = Time.parse('2023-01-02 15:04:05 UTC'); puts t.utc.to_a[0..5].join('-')",
    "5-4-15-2-1-2023"
);
