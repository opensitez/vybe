use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_unboundmethod_creation_instance_method, "class C; def foo; 1; end; end; puts C.instance_method(:foo).class.name", "UnboundMethod");
ruby_test!(test_unboundmethod_creation_unbind, "class C; def foo; 1; end; end; puts C.new.method(:foo).unbind.class.name", "UnboundMethod");
ruby_test!(test_unboundmethod_bind_call, "class C; def foo; 1; end; end; um = C.instance_method(:foo); puts um.bind_call(C.new)", "1");
ruby_test!(test_unboundmethod_bind, "class C; def foo; 1; end; end; um = C.instance_method(:foo); puts um.bind(C.new).call", "1");
ruby_test!(test_unboundmethod_bind_wrong_type, "class C; def foo; 1; end; end; um = C.instance_method(:foo); begin; um.bind(Object.new); rescue TypeError; puts 'err'; end", "err");
ruby_test!(test_unboundmethod_name, "class C; def foo; 1; end; end; puts C.instance_method(:foo).name", "foo");
ruby_test!(test_unboundmethod_owner, "class C; def foo; 1; end; end; puts C.instance_method(:foo).owner", "C");
ruby_test!(test_unboundmethod_arity, "class C; def foo(x); 1; end; end; puts C.instance_method(:foo).arity", "1");
ruby_test!(test_unboundmethod_parameters, "class C; def foo(x, y=1); 1; end; end; puts C.instance_method(:foo).parameters.length", "2");
ruby_test!(test_unboundmethod_super_method, "class A; def foo; 1; end; end; class B < A; def foo; 2; end; end; puts B.instance_method(:foo).super_method.bind_call(B.new)", "1");
