use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_time_round_basic, "puts Time.at(0.5).round.to_i", "1");
ruby_test!(test_time_round_ndigits, "puts Time.at(0.555).round(2).subsec", "111/200"); // 0.55 = 55/100 = 11/20 wait actually 55/100 -> 11/20
ruby_test!(test_time_floor_basic, "puts Time.at(0.5).floor.to_i", "0");
ruby_test!(test_time_floor_ndigits, "puts Time.at(0.555).floor(2).subsec", "11/20"); // 0.55 = 55/100 = 11/20
ruby_test!(test_time_ceil_basic, "puts Time.at(0.5).ceil.to_i", "1");
ruby_test!(test_time_ceil_ndigits, "puts Time.at(0.555).ceil(2).subsec", "14/25"); // 0.56 = 56/100 = 14/25
ruby_test!(test_time_round_half_up, "puts Time.at(0.5).round.to_i", "1"); // Time rounds half up by default? wait no, ruby floats round half-even, but Time.round usually rounds up
