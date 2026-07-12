macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_exception_handling_basic,
    "begin; raise 'err'; rescue; puts 'caught'; end",
    "caught"
);
ruby_test!(
    test_exception_handling_ensure,
    "acc = []; begin; raise 'err'; rescue; acc << 'r'; ensure; acc << 'e'; end; puts acc.join",
    "re"
);
ruby_test!(
    test_exception_handling_else,
    "acc = []; begin; acc << 'b'; rescue; acc << 'r'; else; acc << 'el'; ensure; acc << 'en'; end; puts acc.join",
    "belen"
);
ruby_test!(
    test_exception_handling_else_rescued,
    "acc = []; begin; raise 'err'; rescue; acc << 'r'; else; acc << 'el'; ensure; acc << 'en'; end; puts acc.join",
    "ren"
);
ruby_test!(
    test_exception_handling_retry,
    "acc = 0; begin; acc += 1; raise 'err' if acc < 3; rescue; retry; end; puts acc",
    "3"
);
ruby_test!(
    test_exception_handling_multiple_rescue,
    "begin; raise ArgumentError; rescue TypeError; puts 't'; rescue ArgumentError; puts 'a'; end",
    "a"
);
ruby_test!(
    test_exception_handling_rescue_variable,
    "begin; raise 'err'; rescue => e; puts e.message; end",
    "err"
);
ruby_test!(
    test_exception_handling_rescue_class_variable,
    "begin; raise ArgumentError, 'err'; rescue ArgumentError => e; puts e.message; end",
    "err"
);
ruby_test!(
    test_exception_handling_modifier,
    "puts (raise 'err' rescue 'caught')",
    "caught"
);
ruby_test!(
    test_exception_handling_def_rescue,
    "def foo; raise 'err'; rescue; 'caught'; end; puts foo",
    "caught"
);
