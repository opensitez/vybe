macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_proc_compose_forward,
    "f = proc { |x| x * 2 }; g = proc { |x| x + 1 }; h = f >> g; puts h.call(10)",
    "21"
);
ruby_test!(
    test_proc_compose_backward,
    "f = proc { |x| x * 2 }; g = proc { |x| x + 1 }; h = f << g; puts h.call(10)",
    "22"
);
ruby_test!(
    test_proc_compose_lambda_forward,
    "f = lambda { |x| x * 2 }; g = lambda { |x| x + 1 }; h = f >> g; puts h.call(10)",
    "21"
);
ruby_test!(
    test_proc_compose_lambda_backward,
    "f = lambda { |x| x * 2 }; g = lambda { |x| x + 1 }; h = f << g; puts h.call(10)",
    "22"
);
ruby_test!(
    test_proc_compose_with_method,
    "f = proc { |x| x * 2 }; g = 10.method(:+); h = f >> g; puts h.call(5)",
    "20"
);
