use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_reject_basic, "puts ({a: 1, b: 2, c: 3}.reject {|k, v| v % 2 == 0}.keys.map(&:to_s).join('-'))", "a-c");
ruby_test!(test_reject_returns_hash, "puts ({a: 1}.reject {|k, v| false}.is_a?(Hash))", "true");
ruby_test!(test_reject_does_not_mutate, "h = {a: 1, b: 2}; h.reject {|k, v| true}; puts h.length", "2");
ruby_test!(test_reject_no_block, "puts ({a: 1}.reject.is_a?(Enumerator))", "true");
ruby_test!(test_reject_bang_mutates, "h = {a: 1, b: 2}; h.reject! {|k, v| v == 2}; puts h.keys.map(&:to_s).join('-')", "a");
ruby_test!(test_reject_bang_returns_nil, "h = {a: 1}; puts h.reject! {|k, v| false}.nil?", "true"); // returns nil if no changes made
ruby_test!(test_reject_bang_returns_self, "h = {a: 1}; puts h.reject! {|k, v| true}.object_id == h.object_id", "true"); // returns self if changes made
ruby_test!(test_select_alias, "puts ({a: 1, b: 2}.filter {|k, v| v == 1}.keys.map(&:to_s).join('-'))", "a");
ruby_test!(test_reject_preserves_default, "h = Hash.new('def'); h[:a] = 1; puts h.reject {|k, v| false}.default", "def");
ruby_test!(test_reject_preserves_default_proc, "h = Hash.new {|hash, key| 'def'}; h[:a] = 1; puts h.reject {|k, v| false}.default_proc.is_a?(Proc)", "true");
