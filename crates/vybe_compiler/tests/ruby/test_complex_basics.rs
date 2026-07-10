use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_complex_basic, "puts Complex(1, 2).to_s", "1+2i");
ruby_test!(test_complex_real, "puts Complex(1, 2).real", "1");
ruby_test!(test_complex_imaginary, "puts Complex(1, 2).imaginary", "2");
ruby_test!(test_complex_add, "puts (Complex(1, 2) + Complex(3, 4)).to_s", "4+6i");
ruby_test!(test_complex_sub, "puts (Complex(1, 2) - Complex(3, 4)).to_s", "-2-2i");
ruby_test!(test_complex_mul, "puts (Complex(1, 2) * Complex(3, 4)).to_s", "-5+10i"); // (1*3 - 2*4) + (1*4 + 2*3)i = -5 + 10i
ruby_test!(test_complex_div, "puts (Complex(4, 8) / 2).to_s", "2+4i");
ruby_test!(test_complex_conjugate, "puts Complex(1, 2).conjugate.to_s", "1-2i");
ruby_test!(test_complex_abs, "puts Complex(3, 4).abs", "5.0"); // sqrt(3^2 + 4^2) = 5
