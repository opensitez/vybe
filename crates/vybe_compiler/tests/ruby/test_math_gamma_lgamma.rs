use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_gamma_basic, "puts Math.gamma(5)", "24.0"); // 4!
ruby_test!(test_gamma_half, "puts Math.gamma(0.5).round(5) == Math.sqrt(Math::PI).round(5)", "true");
ruby_test!(test_gamma_zero_domain_error, "begin; Math.gamma(0); rescue Math::DomainError; puts 'err'; end", "err");
ruby_test!(test_gamma_negative_integer_domain_error, "begin; Math.gamma(-1); rescue Math::DomainError; puts 'err'; end", "err");
ruby_test!(test_gamma_negative_float, "puts Math.gamma(-0.5).round(5) == (-2 * Math.sqrt(Math::PI)).round(5)", "true");
ruby_test!(test_lgamma_basic, "puts Math.lgamma(5).map {|x| x.round(5)}.join('-')", "3.17805-1"); // log(24) and sign 1
ruby_test!(test_lgamma_half, "puts Math.lgamma(0.5)[1]", "1");
ruby_test!(test_lgamma_zero_infinity, "puts Math.lgamma(0)[0]", "Infinity");
ruby_test!(test_lgamma_negative_integer_infinity, "puts Math.lgamma(-1)[0]", "Infinity");
ruby_test!(test_lgamma_negative_float, "puts Math.lgamma(-0.5)[1]", "-1");
