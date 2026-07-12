macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_string_comparison_eq, "puts 'hello' == 'hello'", "true");
ruby_test!(
    test_string_comparison_not_eq,
    "puts 'hello' == 'world'",
    "false"
);
ruby_test!(test_string_comparison_lt, "puts 'a' < 'b'", "true");
ruby_test!(test_string_comparison_gt, "puts 'b' > 'a'", "true");
ruby_test!(test_string_comparison_lte, "puts 'a' <= 'a'", "true");
ruby_test!(test_string_comparison_gte, "puts 'b' >= 'a'", "true");
ruby_test!(test_string_comparison_spaceship_eq, "puts 'a' <=> 'a'", "0");
ruby_test!(
    test_string_comparison_spaceship_lt,
    "puts 'a' <=> 'b'",
    "-1"
);
ruby_test!(test_string_comparison_spaceship_gt, "puts 'b' <=> 'a'", "1");
ruby_test!(
    test_string_comparison_spaceship_invalid,
    "puts ('a' <=> 1).nil?",
    "true"
);
ruby_test!(
    test_string_comparison_casecmp,
    "puts 'hello'.casecmp('HELLO')",
    "0"
);
ruby_test!(
    test_string_comparison_casecmp_question,
    "puts 'hello'.casecmp?('HELLO')",
    "true"
);
ruby_test!(
    test_string_comparison_eql,
    "puts 'hello'.eql?('hello')",
    "true"
);
ruby_test!(
    test_string_comparison_include,
    "puts 'hello'.include?('ell')",
    "true"
);
ruby_test!(
    test_string_comparison_start_with,
    "puts 'hello'.start_with?('he')",
    "true"
);
ruby_test!(
    test_string_comparison_end_with,
    "puts 'hello'.end_with?('lo')",
    "true"
);
