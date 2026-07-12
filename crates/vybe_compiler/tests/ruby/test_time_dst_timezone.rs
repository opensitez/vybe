macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_time_now, "puts Time.now.class.name", "Time");
ruby_test!(
    test_time_utc_predicate,
    "puts Time.utc(2000, 1, 1).utc?",
    "true"
);
ruby_test!(
    test_time_utc_predicate_local,
    "puts Time.local(2000, 1, 1).utc?",
    "false"
); // local is false unless local is UTC
ruby_test!(
    test_time_dst_predicate,
    "puts Time.utc(2000, 1, 1).dst?",
    "false"
); // UTC never has DST
ruby_test!(
    test_time_gmt_predicate,
    "puts Time.utc(2000, 1, 1).gmt?",
    "true"
); // alias for utc?
ruby_test!(test_time_zone_utc, "puts Time.utc(2000, 1, 1).zone", "UTC");
ruby_test!(
    test_time_gmtoff_utc,
    "puts Time.utc(2000, 1, 1).gmtoff",
    "0"
);
ruby_test!(test_time_isdst, "puts Time.utc(2000, 1, 1).isdst", "false"); // alias for dst?
