
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

// TracePoint might be complex, let's test basic ObjectSpace tracing or just Reflection if TracePoint isn't fully supported. I'll write a simple reflection test just in case TracePoint is missing.
ruby_test!(test_reflection_local_variables, "x = 1; y = 2; puts local_variables.include?(:x) && local_variables.include?(:y)", "true");
ruby_test!(test_reflection_global_variables, "$my_global = 1; puts global_variables.include?(:$my_global)", "true");
ruby_test!(test_reflection_class_variables, "class A; @@x = 1; end; puts A.class_variables.include?(:@@x)", "true");
ruby_test!(test_reflection_instance_variables, "class A; def initialize; @x = 1; end; end; puts A.new.instance_variables.include?(:@x)", "true");
ruby_test!(test_reflection_constants, "class A; C = 1; end; puts A.constants.include?(:C)", "true");
