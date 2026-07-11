
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_unbound_method_basic, "class A; def foo; 'foo'; end; end; um = A.instance_method(:foo); puts um.class.name", "UnboundMethod");
ruby_test!(test_unbound_method_bind, "class A; def foo; 'foo'; end; end; um = A.instance_method(:foo); m = um.bind(A.new); puts m.call", "foo");
ruby_test!(test_unbound_method_bind_error, "class A; def foo; 'foo'; end; end; class B; end; um = A.instance_method(:foo); begin; um.bind(B.new); rescue TypeError; puts 'err'; end", "err");
ruby_test!(test_unbound_method_name, "class A; def foo; 'foo'; end; end; um = A.instance_method(:foo); puts um.name", "foo");
ruby_test!(test_unbound_method_owner, "class A; def foo; 'foo'; end; end; um = A.instance_method(:foo); puts um.owner == A", "true");
ruby_test!(test_unbound_method_arity, "class A; def foo(x, y); end; end; um = A.instance_method(:foo); puts um.arity", "2");
