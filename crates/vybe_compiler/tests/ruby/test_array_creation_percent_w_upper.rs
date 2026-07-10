use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_percent_W_basic, "puts %W(a b c).join('-')", "a-b-c");
ruby_test!(test_percent_W_spaces, "puts %W(  a   b   c  ).join('-')", "a-b-c");
ruby_test!(test_percent_W_escaped_space, "puts %W(a\\ b c).join('-')", "a b-c");
ruby_test!(test_percent_W_brackets, "puts %W[a b c].join('-')", "a-b-c");
ruby_test!(test_percent_W_braces, "puts %W{a b c}.join('-')", "a-b-c");
ruby_test!(test_percent_W_angles, "puts %W<a b c>.join('-')", "a-b-c");
ruby_test!(test_percent_W_slashes, "puts %W/a b c/.join('-')", "a-b-c");
ruby_test!(test_percent_W_pipes, "puts %W|a b c|.join('-')", "a-b-c");
ruby_test!(test_percent_W_exclamation, "puts %W!a b c!.join('-')", "a-b-c");
ruby_test!(test_percent_W_interp, "x=1; puts %W(a #{x} c).join('-')", "a-1-c");
ruby_test!(test_percent_W_empty, "puts %W().length", "0");
ruby_test!(test_percent_W_escape_seq, "puts %W(a\\nb c).join('-')", "a\nb-c");
