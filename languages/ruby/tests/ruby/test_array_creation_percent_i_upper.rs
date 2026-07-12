macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_percent_I_basic,
    "puts %I(a b c).map{|x| x.class.name}.join('-')",
    "Symbol-Symbol-Symbol"
);
ruby_test!(test_percent_I_values, "puts %I(a b c).join('-')", "a-b-c");
ruby_test!(
    test_percent_I_spaces,
    "puts %I(  a   b   c  ).join('-')",
    "a-b-c"
);
ruby_test!(
    test_percent_I_escaped_space,
    "puts %I(a\\ b c).join('-')",
    "a b-c"
);
ruby_test!(test_percent_I_brackets, "puts %I[a b c].join('-')", "a-b-c");
ruby_test!(test_percent_I_braces, "puts %I{a b c}.join('-')", "a-b-c");
ruby_test!(test_percent_I_angles, "puts %I<a b c>.join('-')", "a-b-c");
ruby_test!(test_percent_I_slashes, "puts %I/a b c/.join('-')", "a-b-c");
ruby_test!(test_percent_I_pipes, "puts %I|a b c|.join('-')", "a-b-c");
ruby_test!(
    test_percent_I_exclamation,
    "puts %I!a b c!.join('-')",
    "a-b-c"
);
ruby_test!(
    test_percent_I_interp,
    "x=1; puts %I(a #{x} c).join('-')",
    "a-1-c"
);
ruby_test!(test_percent_I_empty, "puts %I().length", "0");
