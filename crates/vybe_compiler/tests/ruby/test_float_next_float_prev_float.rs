macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_next_float_basic, "puts 1.0.next_float > 1.0", "true");
ruby_test!(
    test_next_float_infinity,
    "puts Float::INFINITY.next_float == Float::INFINITY",
    "true"
);
ruby_test!(
    test_next_float_nan,
    "puts Float::NAN.next_float.nan?",
    "true"
);
ruby_test!(test_prev_float_basic, "puts 1.0.prev_float < 1.0", "true");
ruby_test!(
    test_prev_float_negative_infinity,
    "puts (-Float::INFINITY).prev_float == -Float::INFINITY",
    "true"
);
ruby_test!(
    test_prev_float_nan,
    "puts Float::NAN.prev_float.nan?",
    "true"
);
ruby_test!(
    test_next_float_prev_float,
    "puts 1.0.next_float.prev_float == 1.0",
    "true"
);
ruby_test!(test_next_float_zero, "puts 0.0.next_float > 0.0", "true");
ruby_test!(test_prev_float_zero, "puts 0.0.prev_float < 0.0", "true");
