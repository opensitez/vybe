
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_percent_q_basic, "puts %Q(hello)", "hello");
ruby_test!(test_percent_q_implicit, "puts %(hello)", "hello");
ruby_test!(test_percent_q_brackets, "puts %Q[hello]", "hello");
ruby_test!(test_percent_q_braces, "puts %Q{hello}", "hello");
ruby_test!(test_percent_q_angles, "puts %Q<hello>", "hello");
ruby_test!(test_percent_q_slashes, "puts %Q/hello/", "hello");
ruby_test!(test_percent_q_pipes, "puts %Q|hello|", "hello");
ruby_test!(test_percent_q_exclamation, "puts %Q!hello!", "hello");
ruby_test!(test_percent_q_interpolation, "name = 'x'; puts %Q(hello #{name})", "hello x");
ruby_test!(test_percent_q_nested_parens, "puts %Q(a (b) c)", "a (b) c");
ruby_test!(test_percent_q_escape_sequences, "puts %Q(a\\nb)", "a\nb");
ruby_test!(test_percent_q_multiline_interp, "x = 1; puts %Q(\n#{x}\n)", "\n1\n");
