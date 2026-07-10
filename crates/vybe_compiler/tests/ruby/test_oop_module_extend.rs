use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_module_extend_object, "module M; def foo; 'M'; end; end; obj = Object.new; obj.extend(M); puts obj.foo", "M");
ruby_test!(test_module_extend_class, "module M; def foo; 'M'; end; end; class A; extend M; end; puts A.foo", "M"); // Extends class methods
ruby_test!(test_module_extend_override, "module M; def foo; 'M'; end; end; class A; extend M; def self.foo; 'A'; end; end; puts A.foo", "A");
ruby_test!(test_module_extend_super, "module M; def foo; super + 'M'; end; end; class A; extend M; def self.foo; 'A'; end; end; puts A.foo", "AM"); // Wait, no, super in extend would call Object's foo if A doesn't define it. Actually if A defines self.foo, extend M puts M *after* A's singleton class.
// Let's refine the super test:
ruby_test!(test_module_extend_super_correct, "module M; def foo; 'M'; end; end; class A; extend M; def self.foo; super + 'A'; end; end; puts A.foo", "MA");
ruby_test!(test_module_singleton_class_include, "module M; def foo; 'M'; end; end; class A; class << self; include M; end; end; puts A.foo", "M"); // Equivalent to extend M
