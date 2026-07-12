macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_method_basic,
    "class A; def foo; 'foo'; end; end; m = A.new.method(:foo); puts m.call",
    "foo"
);
ruby_test!(
    test_method_name,
    "class A; def foo; 'foo'; end; end; m = A.new.method(:foo); puts m.name",
    "foo"
);
ruby_test!(
    test_method_receiver,
    "class A; def foo; 'foo'; end; end; a = A.new; m = a.method(:foo); puts m.receiver == a",
    "true"
);
ruby_test!(
    test_method_owner,
    "class A; def foo; 'foo'; end; end; m = A.new.method(:foo); puts m.owner == A",
    "true"
);
ruby_test!(
    test_method_arity,
    "class A; def foo(x, y); end; end; m = A.new.method(:foo); puts m.arity",
    "2"
);
ruby_test!(
    test_method_to_proc,
    "class A; def foo(x); \"foo_#{x}\"; end; end; m = A.new.method(:foo); p = m.to_proc; puts p.call(1)",
    "foo_1"
);
ruby_test!(
    test_method_unbind,
    "class A; def foo; 'foo'; end; end; m = A.new.method(:foo); um = m.unbind; puts um.class.name",
    "UnboundMethod"
);
