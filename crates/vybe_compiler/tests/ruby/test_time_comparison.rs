use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_time_comparison_eq, "puts Time.utc(2024, 1, 1) == Time.utc(2024, 1, 1)", "true");
ruby_test!(test_time_comparison_not_eq, "puts Time.utc(2024, 1, 1) == Time.utc(2024, 1, 2)", "false");
ruby_test!(test_time_comparison_lt, "puts Time.utc(2024, 1, 1) < Time.utc(2024, 1, 2)", "true");
ruby_test!(test_time_comparison_gt, "puts Time.utc(2024, 1, 2) > Time.utc(2024, 1, 1)", "true");
ruby_test!(test_time_comparison_lte, "puts Time.utc(2024, 1, 1) <= Time.utc(2024, 1, 1)", "true");
ruby_test!(test_time_comparison_gte, "puts Time.utc(2024, 1, 1) >= Time.utc(2024, 1, 1)", "true");
ruby_test!(test_time_comparison_spaceship_eq, "puts Time.utc(2024, 1, 1) <=> Time.utc(2024, 1, 1)", "0");
ruby_test!(test_time_comparison_spaceship_lt, "puts Time.utc(2024, 1, 1) <=> Time.utc(2024, 1, 2)", "-1");
ruby_test!(test_time_comparison_spaceship_gt, "puts Time.utc(2024, 1, 2) <=> Time.utc(2024, 1, 1)", "1");
ruby_test!(test_time_comparison_different_zones, "puts Time.utc(2024, 1, 1, 12, 0, 0) == Time.new(2024, 1, 1, 13, 0, 0, '+01:00')", "true");
ruby_test!(test_time_comparison_spaceship_invalid, "puts (Time.utc(2024, 1, 1) <=> 42).nil?", "true");
