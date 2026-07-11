
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_rescue_basic, "begin; raise 'err'; rescue; puts 'rescued'; end", "rescued");
ruby_test!(test_rescue_class, "begin; raise ArgumentError; rescue ArgumentError; puts 'rescued arg'; end", "rescued arg");
ruby_test!(test_rescue_multiple_classes, "begin; raise TypeError; rescue ArgumentError, TypeError; puts 'rescued either'; end", "rescued either");
ruby_test!(test_rescue_variable, "begin; raise 'err'; rescue => e; puts e.message; end", "err");
ruby_test!(test_rescue_hierarchy, "begin; raise ArgumentError; rescue StandardError; puts 'rescued standard'; end", "rescued standard"); // ArgumentError < StandardError
ruby_test!(test_rescue_default_class, "begin; raise Exception; rescue; puts 'standard'; rescue Exception; puts 'exception'; end", "exception"); // rescue without class only catches StandardError
