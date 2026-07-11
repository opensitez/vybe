
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_integer_pred_basic, "puts 10.pred", "9");
ruby_test!(test_integer_pred_negative, "puts (-10).pred", "-11");
ruby_test!(test_integer_pred_zero, "puts 0.pred", "-1");
ruby_test!(test_integer_succ_basic, "puts 10.succ", "11");
ruby_test!(test_integer_succ_negative, "puts (-10).succ", "-9");
ruby_test!(test_integer_succ_zero, "puts 0.succ", "1");
ruby_test!(test_integer_next_basic, "puts 10.next", "11");
ruby_test!(test_integer_next_negative, "puts (-10).next", "-9");
ruby_test!(test_integer_next_zero, "puts 0.next", "1");
