
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_send_basic, "class A; def foo; 'foo'; end; end; puts A.new.send(:foo)", "foo");
ruby_test!(test_send_args, "class A; def foo(x); \"foo_#{x}\"; end; end; puts A.new.send(:foo, 1)", "foo_1");
ruby_test!(test_send_private, "class A; private; def foo; 'foo'; end; end; puts A.new.send(:foo)", "foo"); // send can call private methods
ruby_test!(test_public_send_basic, "class A; def foo; 'foo'; end; end; puts A.new.public_send(:foo)", "foo");
ruby_test!(test_public_send_private_error, "class A; private; def foo; 'foo'; end; end; begin; A.new.public_send(:foo); rescue NoMethodError; puts 'err'; end", "err"); // public_send cannot call private methods
ruby_test!(test_send_string_name, "class A; def foo; 'foo'; end; end; puts A.new.send('foo')", "foo");
