macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_regex_casefold_predicate_true,
    "puts /a/i.casefold?",
    "true"
);
ruby_test!(
    test_regex_casefold_predicate_false,
    "puts /a/.casefold?",
    "false"
);
ruby_test!(
    test_regex_casefold_predicate_multiline_false,
    "puts /a/m.casefold?",
    "false"
);
