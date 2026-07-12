macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_complex_edge_string_arg,
    "puts Complex('1+2i') == Complex(1, 2)",
    "true"
);
ruby_test!(
    test_complex_edge_float_arg,
    "puts Complex(1.0, 2.0) == Complex(1, 2)",
    "true"
);
ruby_test!(
    test_complex_edge_polar,
    "c = Complex.polar(1, 0); puts c.real == 1.0 && c.imag == 0.0",
    "true"
);
ruby_test!(
    test_complex_edge_rect,
    "c = Complex.rect(1, 2); puts c == Complex(1, 2)",
    "true"
);
ruby_test!(
    test_complex_edge_rectangular,
    "c = Complex.rectangular(1, 2); puts c == Complex(1, 2)",
    "true"
);
ruby_test!(
    test_complex_edge_to_c,
    "c = Complex(1, 2); puts c.to_c.equal?(c)",
    "true"
);
ruby_test!(
    test_complex_edge_to_f,
    "begin; Complex(1, 2).to_f; rescue RangeError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_complex_edge_to_i,
    "begin; Complex(1, 2).to_i; rescue RangeError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_complex_edge_hash,
    "puts Complex(1, 2).hash == Complex(1, 2).hash",
    "true"
);
ruby_test!(
    test_complex_edge_eql,
    "puts Complex(1, 2).eql?(Complex(1, 2))",
    "true"
);
