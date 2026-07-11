
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_script_error_basic, "begin; raise ScriptError; rescue ScriptError; puts 'caught'; end", "caught");
ruby_test!(test_script_error_inherits_exception, "begin; raise ScriptError; rescue Exception; puts 'caught exception'; end", "caught exception");
ruby_test!(test_script_error_not_caught_by_standard_error, "begin; raise ScriptError; rescue StandardError; puts 'caught std'; rescue ScriptError; puts 'caught script'; end", "caught script");
ruby_test!(test_syntax_error_inherits_script_error, "begin; raise SyntaxError; rescue ScriptError; puts 'caught script'; end", "caught script");
ruby_test!(test_load_error_inherits_script_error, "begin; raise LoadError; rescue ScriptError; puts 'caught script'; end", "caught script");
ruby_test!(test_not_implemented_error_inherits_script_error, "begin; raise NotImplementedError; rescue ScriptError; puts 'caught script'; end", "caught script");
