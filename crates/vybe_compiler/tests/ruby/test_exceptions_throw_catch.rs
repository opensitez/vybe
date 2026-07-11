
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_catch_throw_basic, "puts catch(:foo) { throw :foo, 'thrown' }", "thrown");
ruby_test!(test_catch_no_throw, "puts catch(:foo) { 'normal' }", "normal");
ruby_test!(test_catch_throw_nested, "puts catch(:outer) { catch(:inner) { throw :outer, 'out' }; 'in' }", "out");
ruby_test!(test_catch_throw_cross_method, "def bar; throw :foo, 'cross'; end; puts catch(:foo) { bar; 'normal' }", "cross");
ruby_test!(test_catch_throw_uncaught_error, "begin; throw :foo; rescue UncaughtThrowError => e; puts e.tag; end", "foo");
ruby_test!(test_catch_throw_default_value, "puts catch(:foo) { throw :foo } == nil", "true");
