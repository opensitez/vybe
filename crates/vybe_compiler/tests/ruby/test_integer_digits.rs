
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_integer_digits_basic, "puts 12345.digits.join('-')", "5-4-3-2-1");
ruby_test!(test_integer_digits_base_2, "puts 10.digits(2).join('-')", "0-1-0-1");
ruby_test!(test_integer_digits_base_16, "puts 255.digits(16).join('-')", "15-15");
ruby_test!(test_integer_digits_zero, "puts 0.digits.join('-')", "0");
ruby_test!(test_integer_digits_negative, "begin; -10.digits; rescue Math::DomainError; puts 'err'; end", "err");
ruby_test!(test_integer_digits_invalid_base, "begin; 10.digits(1); rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_integer_digits_negative_base, "begin; 10.digits(-2); rescue ArgumentError; puts 'err'; end", "err");
