use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_gcd_basic, "puts 2.gcd(2)", "2");
ruby_test!(test_gcd_coprime, "puts 3.gcd(5)", "1");
ruby_test!(test_gcd_zero, "puts 0.gcd(5)", "5");
ruby_test!(test_gcd_negative, "puts -5.gcd(10)", "5"); // gcd is always positive
ruby_test!(test_lcm_basic, "puts 2.lcm(2)", "2");
ruby_test!(test_lcm_coprime, "puts 3.lcm(5)", "15");
ruby_test!(test_lcm_zero, "puts 0.lcm(5)", "0");
ruby_test!(test_lcm_negative, "puts -5.lcm(10)", "10"); // lcm is always positive
ruby_test!(test_gcdlcm_basic, "puts 3.gcdlcm(5).join('-')", "1-15");
ruby_test!(test_gcdlcm_zero, "puts 0.gcdlcm(5).join('-')", "5-0");
