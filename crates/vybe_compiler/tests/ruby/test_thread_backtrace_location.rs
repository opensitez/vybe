macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_thread_backtrace_location_label,
    "def foo; caller_locations(1, 1).first; end; puts foo.label",
    "foo"
);
ruby_test!(
    test_thread_backtrace_location_base_label,
    "def foo; caller_locations(1, 1).first; end; puts foo.base_label",
    "foo"
);
ruby_test!(
    test_thread_backtrace_location_path,
    "def foo; caller_locations(1, 1).first; end; puts foo.path.include?('eval').to_s",
    "true"
);
ruby_test!(
    test_thread_backtrace_location_absolute_path,
    "def foo; caller_locations(1, 1).first; end; puts foo.absolute_path.nil?",
    "true"
);
ruby_test!(
    test_thread_backtrace_location_lineno,
    "def foo; caller_locations(1, 1).first; end; puts foo.lineno > 0",
    "true"
);
ruby_test!(
    test_thread_backtrace_location_inspect,
    "def foo; caller_locations(1, 1).first; end; puts foo.inspect.include?('foo').to_s",
    "true"
);
ruby_test!(
    test_thread_backtrace_location_to_s,
    "def foo; caller_locations(1, 1).first; end; puts foo.to_s.include?('foo').to_s",
    "true"
);
