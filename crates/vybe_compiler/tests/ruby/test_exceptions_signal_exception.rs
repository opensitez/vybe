use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_signal_exception_basic, "begin; raise SignalException.new('INT'); rescue SignalException => e; puts e.signm; end", "SIGINT");
ruby_test!(test_signal_exception_signo, "begin; raise SignalException.new('INT'); rescue SignalException => e; puts e.signo > 0; end", "true");
ruby_test!(test_signal_exception_numeric, "begin; raise SignalException.new(9); rescue SignalException => e; puts e.signm; end", "SIGKILL");
ruby_test!(test_signal_exception_not_caught_by_standard_error, "begin; raise SignalException.new('INT'); rescue StandardError; puts 'caught'; rescue SignalException; puts 'signal'; end", "signal");
