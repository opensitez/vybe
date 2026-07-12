macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

// require bigdecimal if not loaded automatically
ruby_test!(
    test_bigdecimal_creation_string,
    "require 'bigdecimal'; puts BigDecimal('1.23').to_s('F')",
    "1.23"
);
ruby_test!(
    test_bigdecimal_creation_float_warning,
    "require 'bigdecimal'; puts BigDecimal(1.23, 3).to_s('F')",
    "1.23"
); // often creates a warning if used without precision
ruby_test!(
    test_bigdecimal_arithmetic_add,
    "require 'bigdecimal'; puts (BigDecimal('1.2') + BigDecimal('2.3')).to_s('F')",
    "3.5"
);
ruby_test!(
    test_bigdecimal_arithmetic_mul,
    "require 'bigdecimal'; puts (BigDecimal('1.5') * BigDecimal('2.0')).to_s('F')",
    "3.0"
);
ruby_test!(
    test_bigdecimal_arithmetic_div,
    "require 'bigdecimal'; puts (BigDecimal('5.0') / BigDecimal('2.0')).to_s('F')",
    "2.5"
);
ruby_test!(
    test_bigdecimal_precision,
    "require 'bigdecimal'; puts (BigDecimal('1.0') / BigDecimal('3.0')).round(4).to_s('F')",
    "0.3333"
);
ruby_test!(
    test_bigdecimal_to_f,
    "require 'bigdecimal'; puts BigDecimal('1.5').to_f",
    "1.5"
);
ruby_test!(
    test_bigdecimal_to_i,
    "require 'bigdecimal'; puts BigDecimal('1.9').to_i",
    "1"
);
ruby_test!(
    test_bigdecimal_cmp,
    "require 'bigdecimal'; puts (BigDecimal('1.0') <=> BigDecimal('1.00'))",
    "0"
); // numerically equal
