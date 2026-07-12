macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_interrupt_basic,
    "begin; raise Interrupt; rescue Interrupt; puts 'caught'; end",
    "caught"
);
ruby_test!(
    test_interrupt_inherits_signal_exception,
    "begin; raise Interrupt; rescue SignalException; puts 'caught signal'; end",
    "caught signal"
);
ruby_test!(
    test_interrupt_not_caught_by_standard_error,
    "begin; raise Interrupt; rescue StandardError; puts 'caught std'; rescue Interrupt; puts 'caught int'; end",
    "caught int"
);
ruby_test!(
    test_interrupt_signm,
    "begin; raise Interrupt; rescue Interrupt => e; puts e.signm; end",
    "SIGINT"
);
