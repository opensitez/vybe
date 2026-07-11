
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_angle_positive, "puts 1.5.angle", "0");
ruby_test!(test_angle_zero, "puts 0.0.angle", "0");
ruby_test!(test_angle_negative, "puts (-1.5).angle == Math::PI", "true");
ruby_test!(test_angle_negative_zero, "puts (-0.0).angle == Math::PI", "true");
ruby_test!(test_angle_nan, "puts Float::NAN.angle.nan?", "true");
ruby_test!(test_phase_alias_positive, "puts 1.5.phase", "0");
ruby_test!(test_phase_alias_negative, "puts (-1.5).phase == Math::PI", "true");
ruby_test!(test_arg_alias_positive, "puts 1.5.arg", "0");
ruby_test!(test_arg_alias_negative, "puts (-1.5).arg == Math::PI", "true");
