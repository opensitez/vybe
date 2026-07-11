
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_delete_if_basic, "h = {a: 1, b: 2, c: 3}; h.delete_if {|k, v| v % 2 == 0}; puts h.keys.map(&:to_s).join('-')", "a-c");
ruby_test!(test_delete_if_returns_self, "h = {a: 1}; puts h.delete_if {|k, v| false}.object_id == h.object_id", "true");
ruby_test!(test_delete_if_no_block, "puts {a: 1}.delete_if.is_a?(Enumerator)", "true");
ruby_test!(test_delete_if_all, "h = {a: 1}; h.delete_if {|k, v| true}; puts h.length", "0");
ruby_test!(test_keep_if_basic, "h = {a: 1, b: 2, c: 3}; h.keep_if {|k, v| v % 2 != 0}; puts h.keys.map(&:to_s).join('-')", "a-c");
ruby_test!(test_keep_if_returns_self, "h = {a: 1}; puts h.keep_if {|k, v| true}.object_id == h.object_id", "true");
ruby_test!(test_keep_if_no_block, "puts {a: 1}.keep_if.is_a?(Enumerator)", "true");
ruby_test!(test_keep_if_none, "h = {a: 1}; h.keep_if {|k, v| false}; puts h.length", "0");
ruby_test!(test_delete_if_frozen, "# frozen_string_literal: true\nh = {a: 1}.freeze; begin; h.delete_if {|k, v| true}; rescue FrozenError; puts 'err'; end", "err");
ruby_test!(test_keep_if_frozen, "# frozen_string_literal: true\nh = {a: 1}.freeze; begin; h.keep_if {|k, v| false}; rescue FrozenError; puts 'err'; end", "err");
