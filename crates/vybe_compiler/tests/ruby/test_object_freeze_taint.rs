
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_object_freeze_basic, "o = Object.new; puts o.freeze == o", "true");
ruby_test!(test_object_frozen_true, "o = Object.new.freeze; puts o.frozen?", "true");
ruby_test!(test_object_frozen_false, "o = Object.new; puts o.frozen?", "false");
ruby_test!(test_object_freeze_error, "o = Object.new.freeze; begin; def o.foo; end; rescue FrozenError; puts 'err'; end", "err"); // frozen objects can't be modified
ruby_test!(test_object_taint_basic, "o = Object.new; puts o.taint == o", "true");
ruby_test!(test_object_tainted_true, "o = Object.new.taint; puts o.tainted?", "true"); // Note: taint is deprecated in ruby 2.7+ and removed in 3.2, but vybe might support it or mock it.
ruby_test!(test_object_untaint_basic, "o = Object.new.taint; o.untaint; puts o.tainted?", "false");
ruby_test!(test_object_trust_untrust, "o = Object.new.untrust; puts o.untrusted?", "true"); // trust/untrust are aliases for untaint/taint
