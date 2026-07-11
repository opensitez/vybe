
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_proc_new_basic, "p = Proc.new { 'foo' }; puts p.call", "foo");
ruby_test!(test_proc_new_args, "p = Proc.new { |x| \"foo_#{x}\" }; puts p.call(1)", "foo_1");
ruby_test!(test_proc_call_alias_brackets, "p = Proc.new { |x| \"foo_#{x}\" }; puts p[1]", "foo_1");
ruby_test!(test_proc_call_alias_triple_eq, "p = Proc.new { |x| \"foo_#{x}\" }; puts (p === 1)", "foo_1");
ruby_test!(test_proc_call_alias_yield, "p = Proc.new { |x| \"foo_#{x}\" }; puts p.yield(1)", "foo_1");
ruby_test!(test_proc_arity, "p = Proc.new { |x, y| }; puts p.arity", "2");
ruby_test!(test_proc_lambda_predicate, "p = Proc.new { }; puts p.lambda?", "false");
ruby_test!(test_proc_to_s, "p = Proc.new { }; puts p.to_s.include?('Proc')", "true");
