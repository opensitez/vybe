use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_hook_included, "module M; @acc = []; def self.included(base); @acc << base; end; def self.acc; @acc; end; end; class A; include M; end; puts M.acc.include?(A)", "true");
ruby_test!(test_hook_extended, "module M; @acc = []; def self.extended(base); @acc << base; end; def self.acc; @acc; end; end; class A; extend M; end; puts M.acc.include?(A)", "true");
ruby_test!(test_hook_prepended, "module M; @acc = []; def self.prepended(base); @acc << base; end; def self.acc; @acc; end; end; class A; prepend M; end; puts M.acc.include?(A)", "true");
ruby_test!(test_hook_inherited, "class A; @acc = []; def self.inherited(base); @acc << base; end; def self.acc; @acc; end; end; class B < A; end; puts A.acc.include?(B)", "true");
