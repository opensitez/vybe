use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_send_basic, "class A; def foo; 'foo'; end; end; puts A.new.send(:foo)", "foo");
ruby_test!(test_send_with_args, "class A; def foo(x); \"foo#{x}\"; end; end; puts A.new.send(:foo, 1)", "foo1");
ruby_test!(test_send_private_method, "class A; private; def foo; 'foo'; end; end; puts A.new.send(:foo)", "foo");
ruby_test!(test_public_send_basic, "class A; def foo; 'foo'; end; end; puts A.new.public_send(:foo)", "foo");
ruby_test!(test_public_send_private_method, "class A; private; def foo; 'foo'; end; end; begin; A.new.public_send(:foo); rescue NoMethodError; puts 'err'; end", "err");
ruby_test!(test_send_block, "class A; def foo; yield; end; end; puts A.new.send(:foo) { 'block' }", "block");
