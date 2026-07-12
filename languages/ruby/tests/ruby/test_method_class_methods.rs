macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_method_creation,
    "def foo; 1; end; puts method(:foo).class.name",
    "Method"
);
ruby_test!(
    test_method_call,
    "def foo; 1; end; puts method(:foo).call",
    "1"
);
ruby_test!(
    test_method_bracket,
    "def foo; 1; end; puts method(:foo)[]",
    "1"
);
ruby_test!(
    test_method_receiver,
    "def foo; 1; end; puts method(:foo).receiver.class.name",
    "Object"
);
ruby_test!(
    test_method_name,
    "def foo; 1; end; puts method(:foo).name",
    "foo"
);
ruby_test!(
    test_method_original_name,
    "def foo; 1; end; alias bar foo; puts method(:bar).original_name",
    "foo"
);
ruby_test!(
    test_method_owner,
    "class C; def foo; 1; end; end; puts C.new.method(:foo).owner",
    "C"
);
ruby_test!(
    test_method_unbind,
    "class C; def foo; 1; end; end; puts C.new.method(:foo).unbind.class.name",
    "UnboundMethod"
);
ruby_test!(
    test_method_arity,
    "def foo(x); 1; end; puts method(:foo).arity",
    "1"
);
ruby_test!(
    test_method_parameters,
    "def foo(x, y=1); 1; end; puts method(:foo).parameters.length",
    "2"
);
ruby_test!(
    test_method_super_method,
    "class A; def foo; 1; end; end; class B < A; def foo; 2; end; end; puts B.new.method(:foo).super_method.call",
    "1"
);
ruby_test!(
    test_method_to_proc,
    "def foo(x); x; end; puts [1, 2].map(&method(:foo)).join('-')",
    "1-2"
);
ruby_test!(
    test_method_composition,
    "def f(x); x*2; end; def g(x); x+1; end; puts (method(:f) << method(:g)).call(1)",
    "4"
);
