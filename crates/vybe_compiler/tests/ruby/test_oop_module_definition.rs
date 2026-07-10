use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_module_definition_basic, "module M; end; puts M.class.name", "Module");
ruby_test!(test_module_definition_methods, "module M; def self.foo; 'foo'; end; end; puts M.foo", "foo");
ruby_test!(test_module_reopening, "module M; def self.foo; 'foo'; end; end; module M; def self.bar; 'bar'; end; end; puts \"#{M.foo}-#{M.bar}\"", "foo-bar");
ruby_test!(test_module_name, "module M; end; puts M.name", "M");
ruby_test!(test_module_new_block, "m = Module.new { def foo; 'foo'; end }; class A; include m; end; puts A.new.foo", "foo"); // Need to assign to a variable to use it, or include it directly
