use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_regex_union_basic, "puts Regexp.union('a', 'b').source", "a|b");
ruby_test!(test_regex_union_array, "puts Regexp.union(['a', 'b']).source", "a|b");
ruby_test!(test_regex_union_regexes, "puts Regexp.union(/a/, /b/).source", "(?-mix:a)|(?-mix:b)");
ruby_test!(test_regex_union_mixed, "puts Regexp.union('a', /b/).source", "a|(?-mix:b)");
ruby_test!(test_regex_union_escapes_strings, "puts Regexp.union('.', '*').source", "\\\\.|\\\\*");
ruby_test!(test_regex_union_empty, "puts Regexp.union().source", "(?!)"); // empty union matches nothing
ruby_test!(test_regex_union_empty_array, "puts Regexp.union([]).source", "(?!)");
