
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_float_round, "puts 1.5.round", "2");
ruby_test!(test_float_round_negative, "puts (-1.5).round", "-2");
ruby_test!(test_float_round_precision, "puts 1.555.round(2)", "1.56");
ruby_test!(test_float_floor, "puts 1.5.floor", "1");
ruby_test!(test_float_floor_negative, "puts (-1.5).floor", "-2");
ruby_test!(test_float_floor_precision, "puts 1.555.floor(2)", "1.55");
ruby_test!(test_float_ceil, "puts 1.5.ceil", "2");
ruby_test!(test_float_ceil_negative, "puts (-1.5).ceil", "-1");
ruby_test!(test_float_ceil_precision, "puts 1.555.ceil(2)", "1.56");
ruby_test!(test_float_truncate, "puts 1.5.truncate", "1");
ruby_test!(test_float_truncate_negative, "puts (-1.5).truncate", "-1");
ruby_test!(test_float_truncate_precision, "puts 1.555.truncate(2)", "1.55");
ruby_test!(test_float_next_float, "puts 1.0.next_float > 1.0", "true");
ruby_test!(test_float_prev_float, "puts 1.0.prev_float < 1.0", "true");
ruby_test!(test_float_nan_check, "puts Float::NAN.nan?", "true");
ruby_test!(test_float_infinite_check, "puts Float::INFINITY.infinite?", "1");
ruby_test!(test_float_finite_check, "puts 1.0.finite?", "true");
