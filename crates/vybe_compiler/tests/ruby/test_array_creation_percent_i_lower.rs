use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_percent_i_basic, "puts %i(a b c).map{|x| x.class.name}.join('-')", "Symbol-Symbol-Symbol");
ruby_test!(test_percent_i_values, "puts %i(a b c).join('-')", "a-b-c");
ruby_test!(test_percent_i_spaces, "puts %i(  a   b   c  ).join('-')", "a-b-c");
ruby_test!(test_percent_i_escaped_space, "puts %i(a\\ b c).join('-')", "a b-c");
ruby_test!(test_percent_i_brackets, "puts %i[a b c].join('-')", "a-b-c");
ruby_test!(test_percent_i_braces, "puts %i{a b c}.join('-')", "a-b-c");
ruby_test!(test_percent_i_angles, "puts %i<a b c>.join('-')", "a-b-c");
ruby_test!(test_percent_i_slashes, "puts %i/a b c/.join('-')", "a-b-c");
ruby_test!(test_percent_i_pipes, "puts %i|a b c|.join('-')", "a-b-c");
ruby_test!(test_percent_i_exclamation, "puts %i!a b c!.join('-')", "a-b-c");
ruby_test!(test_percent_i_no_interp, "x=1; puts %i(a #{x} c).join('-')", "a-\\#{x}-c");
ruby_test!(test_percent_i_empty, "puts %i().length", "0");
