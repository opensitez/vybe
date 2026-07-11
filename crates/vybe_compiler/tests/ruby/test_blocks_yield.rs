
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_block_yield_basic, "def foo; yield; end; puts foo { 'foo' }", "foo");
ruby_test!(test_block_yield_args, "def foo; yield 1; end; puts foo { |x| \"foo_#{x}\" }", "foo_1");
ruby_test!(test_block_yield_multiple_args, "def foo; yield 1, 2; end; puts foo { |x, y| \"#{x}_#{y}\" }", "1_2");
ruby_test!(test_block_given_true, "def foo; block_given?; end; puts foo { }", "true");
ruby_test!(test_block_given_false, "def foo; block_given?; end; puts foo", "false");
ruby_test!(test_block_yield_error, "def foo; yield; end; begin; foo; rescue LocalJumpError; puts 'err'; end", "err");
ruby_test!(test_block_pass, "def foo(&b); b.call; end; puts foo { 'foo' }", "foo");
ruby_test!(test_block_pass_to_yield, "def foo(&b); yield; end; puts foo { 'foo' }", "foo");
