use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_append_features_override, "module M; def self.append_features(base); base.define_method(:foo) { 'foo' }; end; end; class A; include M; end; puts A.new.foo", "foo"); // overriding append_features stops normal inclusion unless super is called, but we manually defined method
ruby_test!(test_extend_object_override, "module M; def self.extend_object(base); base.define_singleton_method(:foo) { 'foo' }; end; end; class A; extend M; end; puts A.foo", "foo");
ruby_test!(test_prepend_features_override, "module M; def self.prepend_features(base); base.define_method(:foo) { 'foo' }; end; end; class A; prepend M; end; puts A.new.foo", "foo");
ruby_test!(test_included_modules_after_custom_append, "module M; def self.append_features(base); end; end; class A; include M; end; puts A.included_modules.include?(M)", "false"); // because super wasn't called
