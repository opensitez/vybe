macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_log_basic, "puts Math.log(Math::E)", "1.0");
ruby_test!(test_log_base, "puts Math.log(8, 2)", "3.0");
ruby_test!(
    test_log_negative_domain_error,
    "begin; Math.log(-1); rescue Math::DomainError; puts 'err'; end",
    "err"
);
ruby_test!(test_log_zero, "puts Math.log(0)", "-Infinity");
ruby_test!(test_log10_basic, "puts Math.log10(100)", "2.0");
ruby_test!(
    test_log10_negative_domain_error,
    "begin; Math.log10(-1); rescue Math::DomainError; puts 'err'; end",
    "err"
);
ruby_test!(test_log10_zero, "puts Math.log10(0)", "-Infinity");
ruby_test!(test_log2_basic, "puts Math.log2(8)", "3.0");
ruby_test!(
    test_log2_negative_domain_error,
    "begin; Math.log2(-1); rescue Math::DomainError; puts 'err'; end",
    "err"
);
ruby_test!(test_log2_zero, "puts Math.log2(0)", "-Infinity");
