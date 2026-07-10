use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_proc_creation_proc, "puts Proc.new { 1 }.class.name", "Proc");
ruby_test!(test_proc_creation_lambda, "puts lambda { 1 }.class.name", "Proc");
ruby_test!(test_proc_creation_stabby, "puts (-> { 1 }).class.name", "Proc");
ruby_test!(test_proc_creation_method, "def foo; 1; end; puts method(:foo).to_proc.class.name", "Proc");
ruby_test!(test_proc_creation_symbol, "puts :to_s.to_proc.class.name", "Proc");
ruby_test!(test_proc_call_call, "puts Proc.new { 1 }.call", "1");
ruby_test!(test_proc_call_bracket, "puts Proc.new { 1 }[]", "1");
ruby_test!(test_proc_call_yield, "puts Proc.new { 1 }.yield", "1");
ruby_test!(test_proc_call_dot_call, "puts (->(x){x}).(1)", "1");
ruby_test!(test_proc_call_triple_eq, "puts Proc.new { |x| x.even? } === 2", "true");
ruby_test!(test_proc_lambda_question_proc, "puts Proc.new { 1 }.lambda?", "false");
ruby_test!(test_proc_lambda_question_lambda, "puts lambda { 1 }.lambda?", "true");
ruby_test!(test_proc_lambda_question_stabby, "puts (-> { 1 }).lambda?", "true");
ruby_test!(test_proc_binding, "a = 1; puts Proc.new { a }.binding.class.name", "Binding");
ruby_test!(test_proc_arity, "puts Proc.new { |x, y| x }.arity", "2");
ruby_test!(test_proc_arity_splat, "puts Proc.new { |*x| x }.arity", "-1");
ruby_test!(test_proc_curry, "puts Proc.new { |x, y| x + y }.curry[1][2]", "3");
ruby_test!(test_proc_composition_compose, "f = ->x{x*2}; g = ->x{x+1}; puts (f << g).call(1)", "4");
ruby_test!(test_proc_composition_compose_reverse, "f = ->x{x*2}; g = ->x{x+1}; puts (f >> g).call(1)", "3");
ruby_test!(test_proc_parameters, "puts Proc.new { |x, y=1, *z| }.parameters.length", "3");
