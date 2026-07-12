macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_regexp_options_ignorecase,
    "puts /a/i.match?('A')",
    "true"
);
ruby_test!(
    test_regexp_options_ignorecase_new,
    "puts Regexp.new('a', Regexp::IGNORECASE).match?('A')",
    "true"
);
ruby_test!(
    test_regexp_options_extended,
    "puts /a b/x.match?('ab')",
    "true"
);
ruby_test!(
    test_regexp_options_extended_new,
    "puts Regexp.new('a b', Regexp::EXTENDED).match?('ab')",
    "true"
);
ruby_test!(
    test_regexp_options_multiline,
    "puts /a.*b/m.match?(\"a\\nxb\")",
    "true"
);
ruby_test!(
    test_regexp_options_multiline_new,
    "puts Regexp.new('a.*b', Regexp::MULTILINE).match?(\"a\\nxb\")",
    "true"
);
ruby_test!(
    test_regexp_options_combined,
    "puts /a b/ix.match?('A B')",
    "true"
);
ruby_test!(test_regexp_options_inspect, "puts /a/i.inspect", "/a/i");
ruby_test!(test_regexp_options_to_s, "puts /a/i.to_s", "(?i-mx:a)");
ruby_test!(
    test_regexp_options_options_method,
    "puts (/a/i.options & Regexp::IGNORECASE) > 0",
    "true"
);
