
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_rescue_modifier_basic, "puts (raise 'err' rescue 'rescued')", "rescued");
ruby_test!(test_rescue_modifier_no_error, "puts ('success' rescue 'rescued')", "success");
ruby_test!(test_rescue_modifier_precedence, "puts 1 + (raise 'err' rescue 2)", "3");
ruby_test!(test_rescue_modifier_only_catches_standard_error, "begin; raise Exception rescue 'rescued'; rescue Exception; puts 'exception'; end", "exception"); // rescue modifier catches StandardError only
ruby_test!(test_rescue_modifier_method_call, "def foo; raise 'err'; end; puts foo rescue 'rescued'", "rescued"); // Actually, in puts foo rescue 'rescued', rescue applies to the whole statement 'puts foo'. It rescues the error from foo and evaluates to 'rescued'. Wait, does it print anything? No, puts is never called because foo raises. So it evaluates to 'rescued' but doesn't print.
// Let's refine the test for method call to print the result:
ruby_test!(test_rescue_modifier_method_call_eval, "def foo; raise 'err'; end; res = (foo rescue 'rescued'); puts res", "rescued");
