
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_delete_basic, "h = {a: 1, b: 2}; puts h.delete(:a)", "1");
ruby_test!(test_delete_mutates, "h = {a: 1, b: 2}; h.delete(:a); puts h.keys.map(&:to_s).join('-')", "b");
ruby_test!(test_delete_missing, "h = {a: 1}; puts h.delete(:b).nil?", "true");
ruby_test!(test_delete_missing_with_block, "h = {a: 1}; puts h.delete(:b) {|k| \"def_#{k}\"}", "def_b");
ruby_test!(test_delete_found_ignores_block, "h = {a: 1}; puts h.delete(:a) {|k| 'def'}", "1");
ruby_test!(test_delete_ignores_default, "h = Hash.new('def'); h[:a] = 1; puts h.delete(:b).nil?", "true");
ruby_test!(test_delete_frozen_error, "# frozen_string_literal: true\nh = {a: 1}.freeze; begin; h.delete(:a); rescue FrozenError; puts 'err'; end", "err");
ruby_test!(test_delete_missing_frozen_error, "# frozen_string_literal: true\nh = {a: 1}.freeze; begin; h.delete(:b); rescue FrozenError; puts 'err'; end", "err");
ruby_test!(test_delete_nil_value, "h = {a: nil}; puts h.delete(:a).nil?", "true");
ruby_test!(test_delete_returns_nil_for_nil_value_when_found, "h = {a: nil}; h.delete(:a); puts h.length", "0");
