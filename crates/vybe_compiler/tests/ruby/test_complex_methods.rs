
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_complex_creation, "puts Complex(1, 2).class.name", "Complex");
ruby_test!(test_complex_real, "puts Complex(1, 2).real", "1");
ruby_test!(test_complex_imaginary, "puts Complex(1, 2).imaginary", "2");
ruby_test!(test_complex_addition, "puts (Complex(1, 2) + Complex(3, 4))", "4+6i");
ruby_test!(test_complex_subtraction, "puts (Complex(3, 4) - Complex(1, 2))", "2+2i");
ruby_test!(test_complex_multiplication, "puts (Complex(1, 2) * Complex(3, 4))", "-5+10i");
ruby_test!(test_complex_division, "puts (Complex(1, 2) / Complex(1, 2))", "1+0i");
ruby_test!(test_complex_conjugate, "puts Complex(1, 2).conjugate", "1-2i");
ruby_test!(test_complex_abs, "puts Complex(3, 4).abs", "5.0");
ruby_test!(test_complex_arg, "puts Complex(0, 1).arg.round(2)", "1.57");
ruby_test!(test_complex_polar, "puts Complex(3, 4).polar.class.name", "Array");
ruby_test!(test_complex_rect, "puts Complex(3, 4).rect.join('-')", "3-4");
ruby_test!(test_complex_to_s, "puts Complex(1, -2).to_s", "1-2i");
ruby_test!(test_complex_compare, "puts Complex(1, 2) == Complex(1, 2)", "true");
