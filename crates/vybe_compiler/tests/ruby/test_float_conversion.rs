
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_float_conversion_to_i, "puts 1.5.to_i", "1");
ruby_test!(test_float_conversion_to_i_negative, "puts (-1.5).to_i", "-1");
ruby_test!(test_float_conversion_to_f, "puts 1.5.to_f == 1.5", "true");
ruby_test!(test_float_conversion_to_s, "puts 1.5.to_s", "1.5");
ruby_test!(test_float_conversion_to_r, "puts 1.5.to_r == Rational(3, 2)", "true");
ruby_test!(test_float_conversion_to_c, "puts 1.5.to_c == Complex(1.5, 0)", "true");
ruby_test!(test_float_conversion_truncate, "puts 1.5.truncate", "1");
ruby_test!(test_float_conversion_truncate_negative, "puts (-1.5).truncate", "-1");
ruby_test!(test_float_conversion_round, "puts 1.5.round", "2");
ruby_test!(test_float_conversion_round_negative, "puts (-1.5).round", "-2");
ruby_test!(test_float_conversion_ceil, "puts 1.5.ceil", "2");
ruby_test!(test_float_conversion_ceil_negative, "puts (-1.5).ceil", "-1");
ruby_test!(test_float_conversion_floor, "puts 1.5.floor", "1");
ruby_test!(test_float_conversion_floor_negative, "puts (-1.5).floor", "-2");
