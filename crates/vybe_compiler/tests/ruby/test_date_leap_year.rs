
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_leap_predicate, "require 'date'; puts Date.new(2000, 1, 1).leap?", "true"); // 2000 is leap
ruby_test!(test_leap_predicate_false, "require 'date'; puts Date.new(1900, 1, 1).leap?", "false"); // 1900 is not leap
ruby_test!(test_leap_predicate_four, "require 'date'; puts Date.new(2004, 1, 1).leap?", "true"); // 2004 is leap
ruby_test!(test_gregorian_leap, "require 'date'; puts Date.gregorian_leap?(2000)", "true");
ruby_test!(test_julian_leap, "require 'date'; puts Date.julian_leap?(1900)", "true"); // Julian 1900 is leap
ruby_test!(test_leap_year_class_method, "require 'date'; puts Date.leap?(2000)", "true");
