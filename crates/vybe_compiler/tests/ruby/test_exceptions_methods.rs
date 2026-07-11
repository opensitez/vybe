
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_exception_message, "begin; raise 'err'; rescue => e; puts e.message; end", "err");
ruby_test!(test_exception_to_s, "begin; raise 'err'; rescue => e; puts e.to_s; end", "err");
ruby_test!(test_exception_backtrace, "def foo; raise 'err'; end; begin; foo; rescue => e; puts e.backtrace.is_a?(Array) && e.backtrace.size > 0; end", "true");
ruby_test!(test_exception_cause_nil, "begin; raise 'err'; rescue => e; puts e.cause.nil?; end", "true");
ruby_test!(test_exception_cause_chained, "begin; begin; raise 'err1'; rescue; raise 'err2'; end; rescue => e; puts e.cause.message; end", "err1");
ruby_test!(test_exception_full_message, "begin; raise 'err'; rescue => e; puts e.full_message.include?('err'); end", "true");
ruby_test!(test_exception_set_backtrace, "e = StandardError.new; e.set_backtrace(['line1', 'line2']); puts e.backtrace.join('-')", "line1-line2");
