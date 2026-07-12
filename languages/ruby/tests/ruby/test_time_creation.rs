macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_time_creation_now, "puts Time.now.class.name", "Time");
ruby_test!(
    test_time_creation_new,
    "puts Time.new(2024, 1, 1).year",
    "2024"
);
ruby_test!(
    test_time_creation_new_parts,
    "t = Time.new(2024, 2, 29, 12, 30, 45); puts \"#{t.year}-#{t.month}-#{t.day}-#{t.hour}-#{t.min}-#{t.sec}\"",
    "2024-2-29-12-30-45"
);
ruby_test!(
    test_time_creation_utc,
    "puts Time.utc(2024, 1, 1).utc?",
    "true"
);
ruby_test!(
    test_time_creation_local,
    "puts Time.local(2024, 1, 1).utc?",
    "false"
);
ruby_test!(test_time_creation_at, "puts Time.at(0).utc.year", "1970");
ruby_test!(
    test_time_creation_at_microseconds,
    "puts Time.at(0, 500000).usec",
    "500000"
);
ruby_test!(
    test_time_creation_mktime,
    "puts Time.mktime(2024, 1, 1).year",
    "2024"
);
ruby_test!(
    test_time_creation_invalid_month,
    "begin; Time.new(2024, 13, 1); rescue ArgumentError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_time_creation_invalid_day,
    "begin; Time.new(2024, 2, 30); rescue ArgumentError; puts 'err'; end",
    "err"
);
