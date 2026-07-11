
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_proc_parameters_req, "p = Proc.new { |x| }; puts p.parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "opt:x"); // Proc args are optional by default
ruby_test!(test_lambda_parameters_req, "l = lambda { |x| }; puts l.parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "req:x"); // Lambda args are required
ruby_test!(test_proc_parameters_opt, "p = Proc.new { |x=1| }; puts p.parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "opt:x");
ruby_test!(test_proc_parameters_rest, "p = Proc.new { |*x| }; puts p.parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "rest:x");
ruby_test!(test_proc_parameters_keyreq, "p = Proc.new { |x:| }; puts p.parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "keyreq:x");
ruby_test!(test_proc_parameters_key, "p = Proc.new { |x: 1| }; puts p.parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "key:x");
ruby_test!(test_proc_parameters_keyrest, "p = Proc.new { |**x| }; puts p.parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "keyrest:x");
ruby_test!(test_proc_parameters_block, "p = Proc.new { |&x| }; puts p.parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "block:x");
