
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_time_arithmetic_add, "t = Time.utc(2024, 1, 1); puts (t + 60).min", "1");
ruby_test!(test_time_arithmetic_sub, "t = Time.utc(2024, 1, 1); puts (t - 60).min", "59");
ruby_test!(test_time_arithmetic_diff, "t1 = Time.utc(2024, 1, 1); t2 = Time.utc(2024, 1, 1, 0, 1, 0); puts (t2 - t1).to_i", "60");
ruby_test!(test_time_arithmetic_add_float, "t = Time.utc(2024, 1, 1); puts (t + 1.5).usec", "500000");
ruby_test!(test_time_arithmetic_sub_float, "t = Time.utc(2024, 1, 1, 0, 0, 2); puts (t - 0.5).usec", "500000");
ruby_test!(test_time_arithmetic_diff_float, "t1 = Time.utc(2024, 1, 1); t2 = Time.utc(2024, 1, 1) + 1.5; puts (t2 - t1)", "1.5");
ruby_test!(test_time_arithmetic_add_rational, "t = Time.utc(2024, 1, 1); puts (t + Rational(1, 2)).usec", "500000");
ruby_test!(test_time_arithmetic_invalid_add, "begin; Time.utc(2024) + Time.utc(2025); rescue TypeError; puts 'err'; end", "err");
