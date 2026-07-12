macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_and_basic, "puts (true and false)", "false");
ruby_test!(
    test_and_short_circuit,
    "def foo; puts 'foo'; true; end; puts (false and foo)",
    "false"
); // foo is not printed
ruby_test!(test_or_basic, "puts (false or true)", "true");
ruby_test!(
    test_or_short_circuit,
    "def foo; puts 'foo'; true; end; puts (true or foo)",
    "true"
); // foo is not printed
ruby_test!(test_not_basic, "puts (not true)", "false");
ruby_test!(test_amp_amp_basic, "puts (true && false)", "false");
ruby_test!(test_pipe_pipe_basic, "puts (false || true)", "true");
ruby_test!(test_bang_basic, "puts (!true)", "false");
ruby_test!(
    test_operator_precedence,
    "puts (false and true || true)",
    "false"
); // and has lower precedence than ||
ruby_test!(
    test_operator_precedence_amp,
    "puts (false && true || true)",
    "true"
); // && has higher precedence than ||
