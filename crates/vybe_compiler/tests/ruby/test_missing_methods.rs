use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_method_missing_basic, "class C; def method_missing(m, *args); \"#{m}-#{args.join(',')}\"; end; end; puts C.new.foo(1, 2)", "foo-1,2");
ruby_test!(test_method_missing_super, "class C; def method_missing(m, *args); super; end; end; begin; C.new.foo; rescue NoMethodError; puts 'err'; end", "err");
ruby_test!(test_method_missing_respond_to_missing, "class C; def respond_to_missing?(m, priv); m == :foo; end; end; puts C.new.respond_to?(:foo)", "true");
ruby_test!(test_method_missing_respond_to_missing_method, "class C; def respond_to_missing?(m, priv); m == :foo; end; def method_missing(m, *args); 1; end; end; puts C.new.method(:foo).call", "1");
ruby_test!(test_const_missing_basic, "class C; def self.const_missing(c); \"#{c}\"; end; end; puts C::FOO", "FOO");
ruby_test!(test_const_missing_super, "class C; def self.const_missing(c); super; end; end; begin; C::FOO; rescue NameError; puts 'err'; end", "err");
