
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_return_basic, "def foo; return 1; 2; end; puts foo", "1");
ruby_test!(test_return_implicit, "def foo; 1; end; puts foo", "1");
ruby_test!(test_return_multiple, "def foo; return 1, 2; end; puts foo.join('-')", "1-2"); // returns array
ruby_test!(test_return_block, "def foo; yield; end; def bar; foo { return 'block' }; 'method'; end; puts bar", "block"); // return from block returns from enclosing method
ruby_test!(test_return_error, "begin; eval('return'); rescue SyntaxError; puts 'err'; end", "err"); // return outside method
ruby_test!(test_return_ensure, "def foo; begin; return 1; ensure; return 2; end; end; puts foo", "2");
