macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_exception_methods_message,
    "begin; raise 'err'; rescue => e; puts e.message; end",
    "err"
);
ruby_test!(
    test_exception_methods_backtrace,
    "def foo; raise 'err'; end; begin; foo; rescue => e; puts e.backtrace.class.name; end",
    "Array"
);
ruby_test!(
    test_exception_methods_backtrace_locations,
    "def foo; raise 'err'; end; begin; foo; rescue => e; puts e.backtrace_locations.class.name; end",
    "Array"
);
ruby_test!(
    test_exception_methods_cause,
    "begin; begin; raise 'err1'; rescue; raise 'err2'; end; rescue => e; puts e.cause.message; end",
    "err1"
);
ruby_test!(
    test_exception_methods_full_message,
    "begin; raise 'err'; rescue => e; puts e.full_message.include?('err').to_s; end",
    "true"
);
ruby_test!(
    test_exception_methods_set_backtrace,
    "e = StandardError.new('err'); e.set_backtrace(['a.rb:1']); puts e.backtrace.join",
    "a.rb:1"
);
ruby_test!(
    test_exception_methods_inspect,
    "e = StandardError.new('err'); puts e.inspect",
    "#<StandardError: err>"
);
ruby_test!(
    test_exception_methods_exception,
    "e1 = StandardError.new('err'); e2 = e1.exception('err2'); puts \"#{e1.message}-#{e2.message}\"",
    "err-err2"
);
ruby_test!(
    test_exception_methods_exception_no_arg,
    "e1 = StandardError.new('err'); e2 = e1.exception; puts e1.equal?(e2)",
    "true"
); // returns self if no arg
