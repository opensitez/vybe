macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_no_memory_error_basic,
    "begin; raise NoMemoryError; rescue NoMemoryError; puts 'caught'; end",
    "caught"
);
ruby_test!(
    test_no_memory_error_inherits_exception,
    "begin; raise NoMemoryError; rescue Exception; puts 'caught exception'; end",
    "caught exception"
);
ruby_test!(
    test_no_memory_error_not_caught_by_standard_error,
    "begin; raise NoMemoryError; rescue StandardError; puts 'caught std'; rescue NoMemoryError; puts 'caught nomem'; end",
    "caught nomem"
);
