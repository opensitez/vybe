macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_time_subsec_zero, "puts Time.at(100).subsec", "0");
ruby_test!(
    test_time_subsec_rational,
    "puts Time.at(Rational(1, 2)).subsec",
    "1/2"
);
ruby_test!(test_time_subsec_float, "puts Time.at(0.5).subsec", "1/2"); // stored as rational exact
ruby_test!(test_time_usec_zero, "puts Time.at(100).usec", "0");
ruby_test!(test_time_usec_nonzero, "puts Time.at(0.5).usec", "500000"); // microseconds
ruby_test!(test_time_nsec_zero, "puts Time.at(100).nsec", "0");
ruby_test!(
    test_time_nsec_nonzero,
    "puts Time.at(0.5).nsec",
    "500000000"
); // nanoseconds
ruby_test!(test_time_tv_usec, "puts Time.at(0.5).tv_usec", "500000"); // alias
ruby_test!(test_time_tv_nsec, "puts Time.at(0.5).tv_nsec", "500000000"); // alias
