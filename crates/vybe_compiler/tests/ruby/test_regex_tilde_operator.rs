use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_regex_tilde_basic, "$_ = 'cat'; puts ~ /a/", "1");
ruby_test!(test_regex_tilde_missing, "$_ = 'dog'; puts (~ /a/).nil?", "true");
ruby_test!(test_regex_tilde_nil_dollar_underscore, "$_ = nil; puts (~ /a/).nil?", "true");
ruby_test!(test_regex_tilde_sets_last_match, "$_ = 'cat'; ~ /a/; puts $&", "a");
