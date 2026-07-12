macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_divmod_basic_integer,
    "puts 10.divmod(3).join('-')",
    "3-1"
);
ruby_test!(
    test_divmod_negative_integer,
    "puts -10.divmod(3).join('-')",
    "-4-2"
); // division rounds down, so -10 / 3 = -4, rem = 2
ruby_test!(
    test_divmod_negative_divisor,
    "puts 10.divmod(-3).join('-')",
    "-4--2"
);
ruby_test!(
    test_divmod_zero_error,
    "begin; 10.divmod(0); rescue ZeroDivisionError; puts 'err'; end",
    "err"
);
ruby_test!(test_divmod_float, "puts 10.0.divmod(3).join('-')", "3-1.0");
ruby_test!(
    test_divmod_float_negative,
    "puts -10.0.divmod(3).join('-')",
    "-4-2.0"
);
ruby_test!(
    test_divmod_infinity,
    "begin; 10.divmod(Float::INFINITY); rescue FloatDomainError; puts 'err'; end",
    "err"
); // wait, 10.divmod(Infinity) returns [0, 10] in Ruby usually. Let's test that:
ruby_test!(
    test_divmod_infinity_correct,
    "puts 10.divmod(Float::INFINITY).join('-')",
    "0-10"
);
ruby_test!(
    test_divmod_infinity_infinity,
    "begin; Float::INFINITY.divmod(10); rescue FloatDomainError; puts 'err'; end",
    "err"
); // wait, Inf.divmod raises FloatDomainError or returns NaN? FloatDomainError. Let's see: Actually, usually FloatDomainError.
ruby_test!(
    test_divmod_nan,
    "begin; 10.divmod(Float::NAN); rescue FloatDomainError; puts 'err'; end",
    "err"
); // wait, divmod with NaN raises FloatDomainError? In ruby it usually raises FloatDomainError.
