
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_constant_lookup_basic, "class A; C = 'C'; end; puts A::C", "C");
ruby_test!(test_constant_lookup_nested, "class A; class B; C = 'C'; end; end; puts A::B::C", "C");
ruby_test!(test_constant_lookup_lexical, "C = 'top'; class A; puts C; end", "top");
ruby_test!(test_constant_lookup_inheritance, "class A; C = 'C'; end; class B < A; end; puts B::C", "C");
ruby_test!(test_constant_lookup_module, "module M; C = 'C'; end; class A; include M; end; puts A::C", "C");
ruby_test!(test_constant_reassignment_warning, "C = 1; C = 2; puts C", "2"); // raises warning but assigns
ruby_test!(test_constant_missing, "class A; def self.const_missing(n); \"missing #{n}\"; end; end; puts A::C", "missing C");
