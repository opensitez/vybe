use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_proc_binding_basic, "x = 1; p = Proc.new { x }; puts p.binding.eval('x')", "1");
ruby_test!(test_lambda_binding_basic, "x = 1; l = lambda { x }; puts l.binding.eval('x')", "1");
ruby_test!(test_proc_binding_local_variables, "x = 1; p = Proc.new { y = 2; x }; puts p.binding.local_variables.include?(:x)", "true");
ruby_test!(test_proc_binding_receiver, "class A; def foo; Proc.new { }; end; end; a = A.new; p = a.foo; puts p.binding.receiver == a", "true");
