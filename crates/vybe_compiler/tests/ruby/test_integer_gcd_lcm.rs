macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_integer_gcd_basic, "puts 12.gcd(8)", "4");
ruby_test!(test_integer_gcd_prime, "puts 13.gcd(7)", "1");
ruby_test!(test_integer_gcd_zero, "puts 12.gcd(0)", "12");
ruby_test!(test_integer_gcd_negative, "puts (-12).gcd(8)", "4");
ruby_test!(test_integer_lcm_basic, "puts 12.lcm(8)", "24");
ruby_test!(test_integer_lcm_prime, "puts 13.lcm(7)", "91");
ruby_test!(test_integer_lcm_zero, "puts 12.lcm(0)", "0");
ruby_test!(test_integer_lcm_negative, "puts (-12).lcm(8)", "24");
ruby_test!(
    test_integer_gcdlcm_basic,
    "puts 12.gcdlcm(8).join('-')",
    "4-24"
);
ruby_test!(
    test_integer_gcdlcm_negative,
    "puts (-12).gcdlcm(8).join('-')",
    "4-24"
);
ruby_test!(
    test_integer_gcdlcm_zero,
    "puts 12.gcdlcm(0).join('-')",
    "12-0"
);
