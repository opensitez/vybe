use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_shift_basic, "h = {a: 1, b: 2}; puts h.shift.join('-')", "a-1");
ruby_test!(test_shift_mutates, "h = {a: 1, b: 2}; h.shift; puts h.keys.map(&:to_s).join('-')", "b");
ruby_test!(test_shift_empty, "puts {}.shift.nil?", "true");
ruby_test!(test_shift_returns_array, "h = {a: 1}; puts h.shift.is_a?(Array)", "true");
ruby_test!(test_shift_ignores_default, "h = Hash.new('def'); puts h.shift.nil?", "true");
ruby_test!(test_shift_ignores_default_proc, "h = Hash.new {|hash, key| 'def'}; puts h.shift.nil?", "true");
ruby_test!(test_shift_until_empty, "h = {a: 1, b: 2}; h.shift; h.shift; puts h.shift.nil?", "true");
ruby_test!(test_shift_frozen_error, "# frozen_string_literal: true\nh = {a: 1}.freeze; begin; h.shift; rescue FrozenError; puts 'err'; end", "err");
ruby_test!(test_shift_empty_frozen_error, "# frozen_string_literal: true\nh = {}.freeze; begin; h.shift; rescue FrozenError; puts 'err'; end", "err"); // frozen check happens before empty check usually
