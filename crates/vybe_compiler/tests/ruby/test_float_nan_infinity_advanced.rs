use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_nan_predicate, "puts Float::NAN.nan?", "true");
ruby_test!(test_nan_predicate_false, "puts 1.0.nan?", "false");
ruby_test!(test_infinity_predicate, "puts Float::INFINITY.infinite?", "1");
ruby_test!(test_negative_infinity_predicate, "puts (-Float::INFINITY).infinite?", "-1");
ruby_test!(test_infinity_predicate_false, "puts 1.0.infinite?.nil?", "true"); // infinite? returns nil for finite numbers
ruby_test!(test_finite_predicate, "puts 1.0.finite?", "true");
ruby_test!(test_finite_predicate_infinity, "puts Float::INFINITY.finite?", "false");
ruby_test!(test_finite_predicate_nan, "puts Float::NAN.finite?", "false");
ruby_test!(test_nan_equality, "puts (Float::NAN == Float::NAN)", "false"); // NaN != NaN
ruby_test!(test_infinity_equality, "puts (Float::INFINITY == Float::INFINITY)", "true");
ruby_test!(test_infinity_arithmetic, "puts (Float::INFINITY + 1 == Float::INFINITY)", "true");
ruby_test!(test_infinity_division, "puts (1.0 / Float::INFINITY)", "0.0");
