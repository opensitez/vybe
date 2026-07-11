
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_undef_basic, "class A; def foo; 'foo'; end; undef foo; end; begin; A.new.foo; rescue NoMethodError; puts 'err'; end", "err");
ruby_test!(test_undef_method_method, "class A; def foo; 'foo'; end; undef_method :foo; end; begin; A.new.foo; rescue NoMethodError; puts 'err'; end", "err");
ruby_test!(test_remove_method_basic, "class A; def foo; 'A'; end; end; class B < A; def foo; 'B'; end; remove_method :foo; end; puts B.new.foo", "A"); // remove_method falls back to super
ruby_test!(test_undef_blocks_super, "class A; def foo; 'A'; end; end; class B < A; undef foo; end; begin; B.new.foo; rescue NoMethodError; puts 'err'; end", "err"); // undef blocks superclass resolution
