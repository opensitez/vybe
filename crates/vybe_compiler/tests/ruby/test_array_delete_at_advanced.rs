use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_delete_at_basic, "a = [1, 2, 3]; puts a.delete_at(1)", "2");
ruby_test!(test_delete_at_mutates, "a = [1, 2, 3]; a.delete_at(1); puts a.join('-')", "1-3");
ruby_test!(test_delete_at_negative, "a = [1, 2, 3]; puts a.delete_at(-1)", "3");
ruby_test!(test_delete_at_negative_mutates, "a = [1, 2, 3]; a.delete_at(-2); puts a.join('-')", "1-3");
ruby_test!(test_delete_at_out_of_bounds_high, "a = [1]; puts a.delete_at(5).nil?", "true");
ruby_test!(test_delete_at_out_of_bounds_negative, "a = [1]; puts a.delete_at(-5).nil?", "true");
ruby_test!(test_delete_at_empty, "a = []; puts a.delete_at(0).nil?", "true");
ruby_test!(test_delete_at_returns_deleted, "a = [5, 6]; puts a.delete_at(0)", "5");
ruby_test!(test_delete_at_preserves_frozen, "# frozen_string_literal: true\na = [1].freeze; begin; a.delete_at(0); rescue FrozenError; puts 'err'; end", "err");
ruby_test!(test_delete_at_shifts_elements, "a = [1, 2, 3, 4]; a.delete_at(1); puts a[1]", "3");
