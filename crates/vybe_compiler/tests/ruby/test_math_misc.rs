
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_math_misc_pi, "puts Math::PI > 3.14", "true");
ruby_test!(test_math_misc_e, "puts Math::E > 2.71", "true");
ruby_test!(test_math_misc_domain_error, "begin; Math.sqrt(-1); rescue Math::DomainError; puts 'err'; end", "err");
ruby_test!(test_math_misc_type_error, "begin; Math.sqrt('a'); rescue TypeError; puts 'err'; end", "err");
ruby_test!(test_math_misc_extend, "class MyMath; include Math; end; puts MyMath.new.sqrt(9)", "3.0");
