
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_percent_q_basic, "puts %q(hello)", "hello");
ruby_test!(test_percent_q_brackets, "puts %q[hello]", "hello");
ruby_test!(test_percent_q_braces, "puts %q{hello}", "hello");
ruby_test!(test_percent_q_angles, "puts %q<hello>", "hello");
ruby_test!(test_percent_q_slashes, "puts %q/hello/", "hello");
ruby_test!(test_percent_q_pipes, "puts %q|hello|", "hello");
ruby_test!(test_percent_q_exclamation, "puts %q!hello!", "hello");
ruby_test!(test_percent_q_no_interpolation, "name = 'x'; puts %q(hello #{name})", "hello #{name}");
ruby_test!(test_percent_q_nested_parens, "puts %q(a (b) c)", "a (b) c");
ruby_test!(test_percent_q_nested_brackets, "puts %q[a [b] c]", "a [b] c");
ruby_test!(test_percent_q_nested_braces, "puts %q{a {b} c}", "a {b} c");
ruby_test!(test_percent_q_multiline, "puts %q(\nhello\n)", "\nhello\n");
