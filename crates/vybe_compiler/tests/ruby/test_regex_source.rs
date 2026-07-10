use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_regex_source_basic, "puts /a/.source", "a");
ruby_test!(test_regex_source_escapes, "puts /\\./.source", "\\\\."); // internal escapes are preserved
ruby_test!(test_regex_source_does_not_include_options, "puts /a/i.source", "a");
ruby_test!(test_regex_to_s_basic, "puts /a/.to_s", "(?-mix:a)");
ruby_test!(test_regex_to_s_options, "puts /a/i.to_s", "(?i-mx:a)");
ruby_test!(test_regex_inspect_basic, "puts /a/.inspect", "/a/");
ruby_test!(test_regex_inspect_options, "puts /a/i.inspect", "/a/i");
ruby_test!(test_regex_inspect_escapes, "puts /\\//.inspect", "/\\\\//");
