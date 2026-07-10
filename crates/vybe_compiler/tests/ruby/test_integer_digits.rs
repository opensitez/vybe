use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_digits_basic, "puts 123.digits.join('-')", "3-2-1"); // least significant digit first
ruby_test!(test_digits_zero, "puts 0.digits.join('-')", "0");
ruby_test!(test_digits_negative_error, "begin; -123.digits; rescue Math::DomainError; puts 'err'; end", "err");
ruby_test!(test_digits_base_basic, "puts 16.digits(16).join('-')", "0-1");
ruby_test!(test_digits_base_zero_error, "begin; 16.digits(0); rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_digits_base_negative_error, "begin; 16.digits(-2); rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_digits_base_one_error, "begin; 16.digits(1); rescue ArgumentError; puts 'err'; end", "err");
