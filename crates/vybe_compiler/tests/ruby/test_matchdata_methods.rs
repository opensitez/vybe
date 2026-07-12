macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_matchdata_methods_string,
    "m = /b/.match('abc'); puts m.string",
    "abc"
);
ruby_test!(
    test_matchdata_methods_regexp,
    "m = /b/.match('abc'); puts m.regexp.class.name",
    "Regexp"
);
ruby_test!(
    test_matchdata_methods_length,
    "m = /(a)(b)/.match('abc'); puts m.length",
    "3"
); // full match + 2 captures
ruby_test!(
    test_matchdata_methods_size,
    "m = /(a)(b)/.match('abc'); puts m.size",
    "3"
);
ruby_test!(
    test_matchdata_methods_offset_zero,
    "m = /b/.match('abc'); puts m.offset(0).join('-')",
    "1-2"
);
ruby_test!(
    test_matchdata_methods_offset_capture,
    "m = /a(b)c/.match('abc'); puts m.offset(1).join('-')",
    "1-2"
);
ruby_test!(
    test_matchdata_methods_begin,
    "m = /b/.match('abc'); puts m.begin(0)",
    "1"
);
ruby_test!(
    test_matchdata_methods_end,
    "m = /b/.match('abc'); puts m.end(0)",
    "2"
);
ruby_test!(
    test_matchdata_methods_pre_match,
    "m = /b/.match('abc'); puts m.pre_match",
    "a"
);
ruby_test!(
    test_matchdata_methods_post_match,
    "m = /b/.match('abc'); puts m.post_match",
    "c"
);
