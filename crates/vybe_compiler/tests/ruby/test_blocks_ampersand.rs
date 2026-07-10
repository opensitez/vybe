use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_ampersand_proc_basic, "def foo; yield 1; end; p = Proc.new { |x| \"foo_#{x}\" }; puts foo(&p)", "foo_1");
ruby_test!(test_ampersand_symbol_basic, "puts [1, 2, 3].map(&:to_s).join('-')", "1-2-3");
ruby_test!(test_ampersand_symbol_to_proc, "puts :to_s.to_proc.call(1)", "1");
ruby_test!(test_ampersand_custom_to_proc, "class A; def to_proc; Proc.new { 'foo' }; end; end; def bar; yield; end; puts bar(&A.new)", "foo");
ruby_test!(test_ampersand_nil, "def foo; block_given?; end; puts foo(&nil)", "false");
