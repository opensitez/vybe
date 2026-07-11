
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_alias_basic, "class A; def foo; 'foo'; end; alias bar foo; end; puts A.new.bar", "foo");
ruby_test!(test_alias_method, "class A; def foo; 'foo'; end; alias_method :bar, :foo; end; puts A.new.bar", "foo");
ruby_test!(test_alias_global, "$a = 'a'; alias $b $a; $a = 'c'; puts $b", "c"); // alias global variables
ruby_test!(test_alias_preserves_original, "class A; def foo; 'foo'; end; alias bar foo; def foo; 'foo2'; end; end; puts A.new.bar", "foo"); // alias captures current definition
ruby_test!(test_alias_method_preserves_original, "class A; def foo; 'foo'; end; alias_method :bar, :foo; def foo; 'foo2'; end; end; puts A.new.bar", "foo");
