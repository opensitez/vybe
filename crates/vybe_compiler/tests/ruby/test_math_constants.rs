use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_math_constants_pi, "puts Math::PI.class.name", "Float");
ruby_test!(test_math_constants_e, "puts Math::E.class.name", "Float");
ruby_test!(test_float_constants_nan, "puts Float::NAN.nan?", "true");
ruby_test!(test_float_constants_infinity, "puts Float::INFINITY.infinite?", "1");
ruby_test!(test_float_constants_min, "puts Float::MIN > 0", "true");
ruby_test!(test_float_constants_max, "puts Float::MAX > 0", "true");
ruby_test!(test_float_constants_epsilon, "puts Float::EPSILON > 0", "true");
ruby_test!(test_float_constants_dig, "puts Float::DIG.class.name", "Integer");
ruby_test!(test_float_constants_radix, "puts Float::RADIX.class.name", "Integer");
ruby_test!(test_float_constants_mant_dig, "puts Float::MANT_DIG.class.name", "Integer");
ruby_test!(test_float_constants_min_10_exp, "puts Float::MIN_10_EXP.class.name", "Integer");
ruby_test!(test_float_constants_max_10_exp, "puts Float::MAX_10_EXP.class.name", "Integer");
ruby_test!(test_float_constants_min_exp, "puts Float::MIN_EXP.class.name", "Integer");
ruby_test!(test_float_constants_max_exp, "puts Float::MAX_EXP.class.name", "Integer");
