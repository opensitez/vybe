macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_integer_bitwise_and, "puts 12 & 10", "8");
ruby_test!(test_integer_bitwise_or, "puts 12 | 10", "14");
ruby_test!(test_integer_bitwise_xor, "puts 12 ^ 10", "6");
ruby_test!(test_integer_bitwise_not, "puts ~12", "-13");
ruby_test!(test_integer_shift_left, "puts 12 << 2", "48");
ruby_test!(test_integer_shift_right, "puts 12 >> 2", "3");
ruby_test!(test_integer_bitwise_all_bits, "puts (-1 & 10)", "10");
ruby_test!(test_integer_anybits, "puts 12.anybits?(8)", "true");
ruby_test!(test_integer_anybits_false, "puts 12.anybits?(2)", "false");
ruby_test!(test_integer_allbits, "puts 12.allbits?(12)", "true");
ruby_test!(test_integer_allbits_false, "puts 12.allbits?(14)", "false");
ruby_test!(test_integer_nobits, "puts 12.nobits?(3)", "true");
ruby_test!(test_integer_nobits_false, "puts 12.nobits?(4)", "false");
ruby_test!(test_integer_size, "puts 1.size.class.name", "Integer");
ruby_test!(test_integer_bit_length, "puts 12.bit_length", "4");
ruby_test!(test_integer_digits, "puts 123.digits.join('-')", "3-2-1");
ruby_test!(
    test_integer_digits_base,
    "puts 123.digits(2).join('-')",
    "1-1-0-1-1-1-1"
);
