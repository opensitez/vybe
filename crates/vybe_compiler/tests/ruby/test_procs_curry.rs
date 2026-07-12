macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_proc_curry_basic,
    "p = Proc.new { |x, y| \"#{x}_#{y}\" }; c = p.curry; puts c.call(1).call(2)",
    "1_2"
);
ruby_test!(
    test_proc_curry_arity,
    "p = Proc.new { |x, y| \"#{x}_#{y}\" }; c = p.curry(2); puts c.call(1).call(2)",
    "1_2"
);
ruby_test!(
    test_lambda_curry,
    "l = lambda { |x, y| \"#{x}_#{y}\" }; c = l.curry; puts c.call(1).call(2)",
    "1_2"
);
ruby_test!(
    test_method_curry,
    "class A; def foo(x, y); \"#{x}_#{y}\"; end; end; m = A.new.method(:foo); c = m.to_proc.curry; puts c.call(1).call(2)",
    "1_2"
);
