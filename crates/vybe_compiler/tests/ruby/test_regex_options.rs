use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_regex_options_ignorecase, "puts /A/i.match?('a')", "true");
ruby_test!(test_regex_options_multiline, "puts /a.*b/m.match?(\"a\\nb\")", "true");
ruby_test!(test_regex_options_extended, "puts /a b/x.match?('ab')", "true"); // ignores whitespace
ruby_test!(test_regex_options_combined, "puts /A B/xi.match?('ab')", "true");
ruby_test!(test_regex_options_method, "puts /a/i.options & Regexp::IGNORECASE > 0", "true");
ruby_test!(test_regex_options_multiline_method, "puts /a/m.options & Regexp::MULTILINE > 0", "true");
ruby_test!(test_regex_options_extended_method, "puts /a/x.options & Regexp::EXTENDED > 0", "true");
