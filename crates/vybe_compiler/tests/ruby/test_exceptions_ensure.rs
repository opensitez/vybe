
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_ensure_basic, "acc = []; begin; acc << 'b'; ensure; acc << 'e'; end; puts acc.join('-')", "b-e");
ruby_test!(test_ensure_with_rescue, "acc = []; begin; raise 'err'; rescue; acc << 'r'; ensure; acc << 'e'; end; puts acc.join('-')", "r-e");
ruby_test!(test_ensure_return, "def foo; begin; return 'b'; ensure; return 'e'; end; end; puts foo", "e"); // return in ensure overrides return in begin
ruby_test!(test_ensure_unhandled_exception, "acc = []; begin; begin; raise Exception; ensure; acc << 'e'; end; rescue Exception; acc << 'r'; end; puts acc.join('-')", "e-r"); // ensure runs before exception propagates
ruby_test!(test_ensure_raise, "begin; begin; ensure; raise 'e_err'; end; rescue => e; puts e.message; end", "e_err"); // raise in ensure overrides
