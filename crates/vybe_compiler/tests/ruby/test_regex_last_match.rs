macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_regex_last_match_method_basic,
    "/a/ =~ 'cat'; puts Regexp.last_match(0)",
    "a"
);
ruby_test!(
    test_regex_last_match_method_group,
    "/(a)/ =~ 'cat'; puts Regexp.last_match(1)",
    "a"
);
ruby_test!(
    test_regex_last_match_method_missing,
    "/b/ =~ 'cat'; puts Regexp.last_match.nil?",
    "true"
);
ruby_test!(
    test_regex_last_match_method_no_args,
    "/a/ =~ 'cat'; puts Regexp.last_match.class.name",
    "MatchData"
);
ruby_test!(
    test_regex_last_match_thread_local,
    "t = Thread.new { /a/ =~ 'cat'; Regexp.last_match(0) }; puts t.value",
    "a"
);
