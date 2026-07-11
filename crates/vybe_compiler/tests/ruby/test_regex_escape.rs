
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_regex_escape_basic, "puts Regexp.escape('a.b')", "a\\\\.b");
ruby_test!(test_regex_escape_quote_alias, "puts Regexp.quote('a.b')", "a\\\\.b");
ruby_test!(test_regex_escape_all_meta, "puts Regexp.escape('*?+[]{}()|\\\\.^$')", "\\\\*\\\\?\\\\+\\\\[\\\\]\\\\{\\\\}\\\\(\\\\)\\\\|\\\\\\\\\\\\.\\\\^\\\\$");
ruby_test!(test_regex_escape_spaces, "puts Regexp.escape('a b')", "a\\\\ b");
ruby_test!(test_regex_escape_newlines, "puts Regexp.escape(\"a\\nb\").inspect", "\"a\\\\\\nb\"");
