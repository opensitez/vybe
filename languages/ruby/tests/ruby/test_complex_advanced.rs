macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_complex_creation_basic, "puts Complex(1, 2)", "1+2i");
ruby_test!(test_complex_creation_string, "puts Complex('1+2i')", "1+2i");
ruby_test!(
    test_complex_real_imag,
    "c = Complex(3, 4); puts \"#{c.real}-#{c.imaginary}\"",
    "3-4"
);
ruby_test!(
    test_complex_arithmetic_add,
    "puts Complex(1, 2) + Complex(2, 3)",
    "3+5i"
);
ruby_test!(
    test_complex_arithmetic_mul,
    "puts Complex(1, 2) * Complex(2, 3)",
    "-4+7i"
);
ruby_test!(
    test_complex_arithmetic_int,
    "puts Complex(1, 2) + 1",
    "2+2i"
);
ruby_test!(test_complex_conjugate, "puts Complex(1, 2).conj", "1-2i");
ruby_test!(test_complex_abs, "puts Complex(3, 4).abs", "5.0");
ruby_test!(test_complex_abs2, "puts Complex(3, 4).abs2", "25"); // 3^2 + 4^2
ruby_test!(
    test_complex_rect,
    "puts Complex(1, 2).rect.join('-')",
    "1-2"
);
ruby_test!(
    test_complex_polar_creation,
    "puts Complex.polar(5, 0).round(5).to_s",
    "5.0+0.0i"
);
