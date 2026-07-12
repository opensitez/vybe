macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_eval_basic, "x = 1; puts eval('x + 2')", "3");
ruby_test!(
    test_eval_binding,
    "def foo; x = 1; binding; end; puts eval('x + 2', foo)",
    "3"
);
ruby_test!(
    test_eval_string_interpolation,
    "x = 2; puts eval(\"#{x} + 2\")",
    "4"
); // actually string interpolation is resolved before eval
ruby_test!(
    test_eval_local_variable_assignment,
    "eval('x = 5'); puts x rescue puts 'err'",
    "err"
); // eval introduces a new scope for locals unless they exist? Actually no, local variable assignment in eval without binding creates a local var in eval's scope only if it wasn't defined before. Since run_ruby_one puts everything in one string, let's test it:
ruby_test!(
    test_eval_local_assignment_works_if_defined,
    "x = 1; eval('x = 5'); puts x",
    "5"
);
ruby_test!(
    test_eval_with_file_line,
    "puts eval('__LINE__', nil, 'test.rb', 10)",
    "10"
);
