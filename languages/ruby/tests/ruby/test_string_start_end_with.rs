macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_start_with_true,
    "puts 'hello'.start_with?('he')",
    "true"
);
ruby_test!(
    test_start_with_false,
    "puts 'hello'.start_with?('el')",
    "false"
);
ruby_test!(
    test_start_with_multiple_true,
    "puts 'hello'.start_with?('x', 'he')",
    "true"
);
ruby_test!(
    test_start_with_multiple_false,
    "puts 'hello'.start_with?('x', 'y')",
    "false"
);
ruby_test!(
    test_start_with_regex,
    "puts 'hello'.start_with?(/h[aeiou]/)",
    "true"
);
ruby_test!(
    test_start_with_empty,
    "puts 'hello'.start_with?('')",
    "true"
);
ruby_test!(test_end_with_true, "puts 'hello'.end_with?('lo')", "true");
ruby_test!(test_end_with_false, "puts 'hello'.end_with?('ll')", "false");
ruby_test!(
    test_end_with_multiple_true,
    "puts 'hello'.end_with?('x', 'lo')",
    "true"
);
ruby_test!(
    test_end_with_multiple_false,
    "puts 'hello'.end_with?('x', 'y')",
    "false"
);
ruby_test!(test_end_with_empty, "puts 'hello'.end_with?('')", "true");
ruby_test!(
    test_start_end_both,
    "puts 'hello'.start_with?('h') && 'hello'.end_with?('o')",
    "true"
);
