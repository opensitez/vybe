
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_coerce_integer_integer, "puts 1.coerce(2).join('-')", "2-1");
ruby_test!(test_coerce_float_integer, "puts 1.coerce(2.5).join('-')", "2.5-1.0"); // integer coerces to float
ruby_test!(test_coerce_integer_float, "puts 1.5.coerce(2).join('-')", "2.0-1.5"); // float coerces to float
ruby_test!(test_coerce_rational_integer, "puts Rational(1, 2).coerce(2).join('-')", "2/1-1/2");
ruby_test!(test_coerce_complex_integer, "puts Complex(1, 2).coerce(2).join('-')", "2+0i-1+2i");
ruby_test!(test_coerce_string_error, "begin; 1.coerce('a'); rescue TypeError; puts 'err'; end", "err");
ruby_test!(test_coerce_operation_fallback, "class A; def coerce(other); [other, 2]; end; def *(other); other * 3; end; end; puts A.new * 5", "15"); // custom coerce allows operators to work
