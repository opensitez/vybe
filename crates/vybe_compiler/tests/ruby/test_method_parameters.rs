use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_method_parameters_req, "class A; def foo(x); end; end; puts A.new.method(:foo).parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "req:x");
ruby_test!(test_method_parameters_opt, "class A; def foo(x=1); end; end; puts A.new.method(:foo).parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "opt:x");
ruby_test!(test_method_parameters_rest, "class A; def foo(*x); end; end; puts A.new.method(:foo).parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "rest:x");
ruby_test!(test_method_parameters_keyreq, "class A; def foo(x:); end; end; puts A.new.method(:foo).parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "keyreq:x");
ruby_test!(test_method_parameters_key, "class A; def foo(x: 1); end; end; puts A.new.method(:foo).parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "key:x");
ruby_test!(test_method_parameters_keyrest, "class A; def foo(**x); end; end; puts A.new.method(:foo).parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "keyrest:x");
ruby_test!(test_method_parameters_block, "class A; def foo(&x); end; end; puts A.new.method(:foo).parameters.map{|t,n| \"#{t}:#{n}\"}.join('-')", "block:x");
