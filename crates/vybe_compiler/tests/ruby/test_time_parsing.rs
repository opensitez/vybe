use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_time_parsing_require_time, "require 'time'; puts Time.parse('2024-01-01 12:00:00 UTC').year", "2024");
ruby_test!(test_time_parsing_iso8601, "require 'time'; puts Time.iso8601('2024-02-29T12:30:45Z').month", "2");
ruby_test!(test_time_parsing_rfc2822, "require 'time'; puts Time.rfc2822('Mon, 01 Jan 2024 12:00:00 +0000').year", "2024");
ruby_test!(test_time_parsing_httpdate, "require 'time'; puts Time.httpdate('Mon, 01 Jan 2024 12:00:00 GMT').year", "2024");
ruby_test!(test_time_parsing_strptime, "require 'time'; puts Time.strptime('2024-02-29', '%Y-%m-%d').month", "2");
ruby_test!(test_time_parsing_invalid_parse, "require 'time'; begin; Time.parse('invalid date'); rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_time_parsing_invalid_strptime, "require 'time'; begin; Time.strptime('2024', '%Y-%m'); rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_time_parsing_xmlschema, "require 'time'; puts Time.xmlschema('2024-02-29T12:30:45Z').day", "29");
