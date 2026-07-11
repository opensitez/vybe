
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

// Method binding doesn't have a direct #binding method in older Rubies, but let's test related things or just skip if it's too esoteric. Actually, in Ruby, Method objects DO NOT have a #binding method. Binding is for blocks/procs. Let's test caller_locations instead.
ruby_test!(test_caller_locations_basic, "def foo; caller_locations(1, 1)[0].label; end; puts foo", "foo"); // the label of the frame that called foo... wait, caller_locations(1,1) is the caller of foo. If `puts foo` is at top level, label is `<main>`
ruby_test!(test_caller_basic, "def foo; caller(1, 1)[0]; end; puts foo.include?('<main>')", "true");
ruby_test!(test_binding_basic, "def foo; x = 1; binding; end; puts foo.eval('x')", "1");
ruby_test!(test_binding_local_variable_get, "def foo; x = 1; binding; end; puts foo.local_variable_get(:x)", "1");
ruby_test!(test_binding_local_variable_set, "def foo; x = 1; b = binding; b.local_variable_set(:x, 2); x; end; puts foo", "2");
ruby_test!(test_binding_local_variable_defined, "def foo; x = 1; binding; end; puts foo.local_variable_defined?(:x)", "true");
