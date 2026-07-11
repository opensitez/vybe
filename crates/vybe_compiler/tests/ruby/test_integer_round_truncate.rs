
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_round_basic, "puts 10.round", "10");
ruby_test!(test_round_negative_ndigits, "puts 15.round(-1)", "20");
ruby_test!(test_round_negative_ndigits_down, "puts 14.round(-1)", "10");
ruby_test!(test_round_half_up, "puts 15.round(-1, half: :up)", "20");
ruby_test!(test_round_half_down, "puts 15.round(-1, half: :down)", "10");
ruby_test!(test_round_half_even, "puts 15.round(-1, half: :even)", "20");
ruby_test!(test_round_half_even_down, "puts 25.round(-1, half: :even)", "20");
ruby_test!(test_truncate_basic, "puts 10.truncate", "10");
ruby_test!(test_truncate_negative_ndigits, "puts 15.truncate(-1)", "10");
ruby_test!(test_truncate_negative_ndigits_negative_num, "puts -15.truncate(-1)", "-10");
ruby_test!(test_floor_basic, "puts 10.floor", "10");
ruby_test!(test_floor_negative_ndigits, "puts 15.floor(-1)", "10");
ruby_test!(test_floor_negative_ndigits_negative_num, "puts -15.floor(-1)", "-20");
ruby_test!(test_ceil_basic, "puts 10.ceil", "10");
ruby_test!(test_ceil_negative_ndigits, "puts 15.ceil(-1)", "20");
ruby_test!(test_ceil_negative_ndigits_negative_num, "puts -15.ceil(-1)", "-10");
