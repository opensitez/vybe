
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_math_pi, "puts Math::PI > 3.14 && Math::PI < 3.15", "true");
ruby_test!(test_math_e, "puts Math::E > 2.71 && Math::E < 2.72", "true");
ruby_test!(test_math_domain_error, "begin; Math.sqrt(-1); rescue Math::DomainError; puts 'err'; end", "err");
ruby_test!(test_math_domain_error_log, "begin; Math.log(-1); rescue Math::DomainError; puts 'err'; end", "err");
ruby_test!(test_math_domain_error_acos, "begin; Math.acos(2); rescue Math::DomainError; puts 'err'; end", "err");
