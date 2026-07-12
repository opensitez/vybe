macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_even_basic, "puts 2.even?", "true");
ruby_test!(test_even_false, "puts 3.even?", "false");
ruby_test!(test_even_zero, "puts 0.even?", "true");
ruby_test!(test_even_negative, "puts -2.even?", "true");
ruby_test!(test_odd_basic, "puts 3.odd?", "true");
ruby_test!(test_odd_false, "puts 2.odd?", "false");
ruby_test!(test_odd_negative, "puts -3.odd?", "true");
ruby_test!(test_allbits_basic, "puts 0b1010.allbits?(0b1000)", "true");
ruby_test!(test_allbits_false, "puts 0b1010.allbits?(0b1100)", "false");
ruby_test!(test_anybits_basic, "puts 0b1010.anybits?(0b1100)", "true");
ruby_test!(test_anybits_false, "puts 0b1010.anybits?(0b0101)", "false");
ruby_test!(test_nobits_basic, "puts 0b1010.nobits?(0b0101)", "true");
ruby_test!(test_nobits_false, "puts 0b1010.nobits?(0b1000)", "false");
