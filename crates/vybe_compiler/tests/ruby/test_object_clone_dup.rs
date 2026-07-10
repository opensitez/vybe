use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_object_clone_basic, "class A; attr_accessor :x; end; a = A.new; a.x = 1; b = a.clone; puts b.x", "1");
ruby_test!(test_object_dup_basic, "class A; attr_accessor :x; end; a = A.new; a.x = 1; b = a.dup; puts b.x", "1");
ruby_test!(test_object_clone_freeze, "a = Object.new.freeze; b = a.clone; puts b.frozen?", "true"); // clone copies frozen state
ruby_test!(test_object_dup_freeze, "a = Object.new.freeze; b = a.dup; puts b.frozen?", "false"); // dup does not copy frozen state
ruby_test!(test_object_clone_singleton, "a = Object.new; def a.foo; 'foo'; end; b = a.clone; puts b.foo", "foo"); // clone copies singleton methods
ruby_test!(test_object_dup_singleton, "a = Object.new; def a.foo; 'foo'; end; b = a.dup; begin; b.foo; rescue NoMethodError; puts 'err'; end", "err"); // dup does not copy singleton methods
ruby_test!(test_object_clone_kwarg_freeze, "a = Object.new.freeze; b = a.clone(freeze: false); puts b.frozen?", "false"); // ruby 2.4+ clone freeze kwarg
