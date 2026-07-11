
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_method_visibility_public_basic, "class C; public; def foo; 1; end; end; puts C.new.foo", "1");
ruby_test!(test_method_visibility_private_basic, "class C; private; def foo; 1; end; end; begin; C.new.foo; rescue NoMethodError; puts 'err'; end", "err");
ruby_test!(test_method_visibility_private_implicit_receiver, "class C; private; def foo; 1; end; public; def bar; foo; end; end; puts C.new.bar", "1");
ruby_test!(test_method_visibility_protected_basic, "class C; protected; def foo; 1; end; end; begin; C.new.foo; rescue NoMethodError; puts 'err'; end", "err");
ruby_test!(test_method_visibility_protected_same_class, "class C; protected; def foo; 1; end; public; def bar(other); other.foo; end; end; puts C.new.bar(C.new)", "1");
ruby_test!(test_method_visibility_private_class_method, "class C; class << self; private; def foo; 1; end; end; end; begin; C.foo; rescue NoMethodError; puts 'err'; end", "err");
ruby_test!(test_method_visibility_private_send, "class C; private; def foo; 1; end; end; puts C.new.send(:foo)", "1");
ruby_test!(test_method_visibility_public_send, "class C; private; def foo; 1; end; end; begin; C.new.public_send(:foo); rescue NoMethodError; puts 'err'; end", "err");
ruby_test!(test_method_visibility_toplevel_private, "def foo; 1; end; begin; self.foo; rescue NoMethodError; puts 'err'; end", "err"); // top-level methods are private
ruby_test!(test_method_visibility_module_function, "module M; module_function; def foo; 1; end; end; begin; class C; include M; end; C.new.foo; rescue NoMethodError; puts 'err'; end", "err"); // module_function makes instance methods private
