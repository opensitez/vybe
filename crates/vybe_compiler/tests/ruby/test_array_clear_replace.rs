use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_clear_basic, "a = [1, 2]; a.clear; puts a.length", "0");
ruby_test!(test_clear_returns_self, "a = [1, 2]; puts a.clear.object_id == a.object_id", "true");
ruby_test!(test_clear_empty, "a = []; a.clear; puts a.length", "0");
ruby_test!(test_replace_basic, "a = [1]; a.replace([2, 3]); puts a.join('-')", "2-3");
ruby_test!(test_replace_returns_self, "a = [1]; puts a.replace([2]).object_id == a.object_id", "true");
ruby_test!(test_replace_changes_length, "a = [1]; a.replace([1, 2, 3, 4]); puts a.length", "4");
ruby_test!(test_replace_with_empty, "a = [1]; a.replace([]); puts a.length", "0");
ruby_test!(test_replace_empty_with_items, "a = []; a.replace([1]); puts a.join('-')", "1");
ruby_test!(test_replace_self, "a = [1]; a.replace(a); puts a.join('-')", "1");
ruby_test!(test_replace_frozen_error, "# frozen_string_literal: true\na = [1].freeze; begin; a.replace([2]); rescue FrozenError; puts 'err'; end", "err");
ruby_test!(test_clear_frozen_error, "# frozen_string_literal: true\na = [1].freeze; begin; a.clear; rescue FrozenError; puts 'err'; end", "err");
