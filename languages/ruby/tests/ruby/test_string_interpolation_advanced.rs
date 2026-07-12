macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_interp_basic, "x = 5; puts \"#{x}\"", "5");
ruby_test!(test_interp_multiple, "a=1; b=2; puts \"#{a}-#{b}\"", "1-2");
ruby_test!(test_interp_math, "puts \"#{2 * 3}\"", "6");
ruby_test!(test_interp_method_call, "puts \"#{\"abc\".upcase}\"", "ABC");
ruby_test!(test_interp_nested, "puts \"#{ \"#{1}\" }\"", "1");
ruby_test!(test_interp_global, "$g=9; puts \"#$g\"", "9");
ruby_test!(test_interp_instance_var, "@i=8; puts \"#@i\"", "8");
ruby_test!(
    test_interp_class_var,
    "class A; @@c=7; def f; puts \"#@@c\"; end; end; A.new.f",
    "7"
);
ruby_test!(
    test_interp_block,
    "puts \"#{[1,2].map { |x| x*2 }.join(',')}\"",
    "2,4"
);
ruby_test!(test_interp_multiline, "puts \"#{\n2+2\n}\"", "4");
ruby_test!(test_interp_empty, "puts \"#{}\"", "");
ruby_test!(test_interp_no_braces_var, "x=3; puts \"#x\"", "#x"); // only #$g, #@i, #@@c work without braces
