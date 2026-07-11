
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_percent_w_basic, "puts %w(a b c).join('-')", "a-b-c");
ruby_test!(test_percent_w_spaces, "puts %w(  a   b   c  ).join('-')", "a-b-c");
ruby_test!(test_percent_w_newline, "puts %w(a\nb\nc).join('-')", "a-b-c");
ruby_test!(test_percent_w_escaped_space, "puts %w(a\\ b c).join('-')", "a b-c");
ruby_test!(test_percent_w_brackets, "puts %w[a b c].join('-')", "a-b-c");
ruby_test!(test_percent_w_braces, "puts %w{a b c}.join('-')", "a-b-c");
ruby_test!(test_percent_w_angles, "puts %w<a b c>.join('-')", "a-b-c");
ruby_test!(test_percent_w_slashes, "puts %w/a b c/.join('-')", "a-b-c");
ruby_test!(test_percent_w_pipes, "puts %w|a b c|.join('-')", "a-b-c");
ruby_test!(test_percent_w_exclamation, "puts %w!a b c!.join('-')", "a-b-c");
ruby_test!(test_percent_w_no_interp, "x=1; puts %w(a #{x} c).join('-')", "a-\\#{x}-c");
ruby_test!(test_percent_w_empty, "puts %w().length", "0");
