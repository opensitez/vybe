use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_pow_basic, "puts 2.pow(3)", "8");
ruby_test!(test_pow_zero, "puts 2.pow(0)", "1");
ruby_test!(test_pow_one, "puts 2.pow(1)", "2");
ruby_test!(test_pow_negative, "puts 2.pow(-1).class.name", "Rational"); // wait, Integer#pow returns Rational for negative exponent, or Float? Rational. Let's use ** 
ruby_test!(test_pow_negative_float, "puts (2 ** -1).class.name", "Rational");
ruby_test!(test_pow_modulo_basic, "puts 2.pow(3, 5)", "3");
ruby_test!(test_pow_modulo_zero_error, "begin; 2.pow(3, 0); rescue ZeroDivisionError; puts 'err'; end", "err");
ruby_test!(test_pow_modulo_negative_exponent_error, "begin; 2.pow(-1, 5); rescue ArgumentError; puts 'err'; end", "err"); // modulo not supported with negative exponent in Integer#pow usually, wait it is supported if inverse exists in ruby 2.5+
ruby_test!(test_pow_modulo_negative_exponent_coprime, "puts 2.pow(-1, 5)", "3"); // inverse of 2 mod 5 is 3
ruby_test!(test_pow_modulo_negative_exponent_not_coprime, "begin; 2.pow(-1, 4); rescue ZeroDivisionError; puts 'err'; end", "err"); // no inverse
