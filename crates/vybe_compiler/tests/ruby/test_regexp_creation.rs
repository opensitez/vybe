use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_regexp_creation_literal, "puts /abc/.class.name", "Regexp");
ruby_test!(test_regexp_creation_new, "puts Regexp.new('abc').class.name", "Regexp");
ruby_test!(test_regexp_creation_percent_r, "puts %r{abc}.class.name", "Regexp");
ruby_test!(test_regexp_creation_compile, "puts Regexp.compile('abc').class.name", "Regexp");
ruby_test!(test_regexp_creation_escape, "puts Regexp.escape('a.b*c?').class.name", "String");
ruby_test!(test_regexp_creation_quote, "puts Regexp.quote('a.b*c?').class.name", "String");
ruby_test!(test_regexp_creation_interpolation, "str = 'abc'; puts /#{str}/.class.name", "Regexp");
ruby_test!(test_regexp_creation_union_strings, "puts Regexp.union('a', 'b', 'c').class.name", "Regexp");
ruby_test!(test_regexp_creation_union_regexps, "puts Regexp.union(/a/, /b/).class.name", "Regexp");
ruby_test!(test_regexp_creation_union_array, "puts Regexp.union(['a', 'b', 'c']).class.name", "Regexp");
