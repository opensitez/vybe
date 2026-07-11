
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_method_super_method_basic, "class A; def foo; 'A'; end; end; class B < A; def foo; 'B'; end; end; m = B.new.method(:foo); puts m.super_method.call", "A");
ruby_test!(test_method_super_method_missing, "class A; def foo; 'A'; end; end; m = A.new.method(:foo); puts m.super_method.nil?", "true");
ruby_test!(test_unbound_super_method_basic, "class A; def foo; 'A'; end; end; class B < A; def foo; 'B'; end; end; um = B.instance_method(:foo); puts um.super_method.owner == A", "true");
ruby_test!(test_unbound_super_method_missing, "class A; def foo; 'A'; end; end; um = A.instance_method(:foo); puts um.super_method.nil?", "true");
