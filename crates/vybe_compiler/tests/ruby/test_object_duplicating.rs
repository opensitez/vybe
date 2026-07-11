
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_object_dup, "a = [1, 2]; b = a.dup; a << 3; puts b.length", "2");
ruby_test!(test_object_dup_frozen, "a = [1, 2].freeze; puts a.dup.frozen?", "false");
ruby_test!(test_object_clone, "a = [1, 2]; b = a.clone; a << 3; puts b.length", "2");
ruby_test!(test_object_clone_frozen, "a = [1, 2].freeze; puts a.clone.frozen?", "true");
ruby_test!(test_object_clone_frozen_kw, "a = [1, 2].freeze; puts a.clone(freeze: false).frozen?", "false");
ruby_test!(test_object_freeze, "a = [1, 2]; a.freeze; puts a.frozen?", "true");
ruby_test!(test_object_freeze_mutation, "a = [1, 2].freeze; begin; a << 3; rescue FrozenError; puts 'err'; end", "err");
ruby_test!(test_object_taint_deprecated, "a = [1, 2]; a.taint; puts a.tainted?", "false"); // Taint is deprecated and a no-op in Ruby 2.7+ / 3.0+
ruby_test!(test_object_untaint_deprecated, "a = [1, 2]; a.untaint; puts a.tainted?", "false");
ruby_test!(test_object_trust_deprecated, "a = [1, 2]; a.trust; puts a.untrusted?", "false");
ruby_test!(test_object_untrust_deprecated, "a = [1, 2]; a.untrust; puts a.untrusted?", "false");
