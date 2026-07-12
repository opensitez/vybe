macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_delete_if_basic,
    "a = [1, 2, 3, 4]; a.delete_if {|x| x % 2 == 0}; puts a.join('-')",
    "1-3"
);
ruby_test!(
    test_delete_if_returns_self,
    "a = [1, 2]; puts a.delete_if {|x| false}.object_id == a.object_id",
    "true"
);
ruby_test!(
    test_delete_if_no_block,
    "puts [1, 2].delete_if.is_a?(Enumerator)",
    "true"
);
ruby_test!(
    test_delete_if_all,
    "a = [1, 2]; a.delete_if {|x| true}; puts a.length",
    "0"
);
ruby_test!(
    test_delete_if_none,
    "a = [1, 2]; a.delete_if {|x| false}; puts a.length",
    "2"
);
ruby_test!(
    test_keep_if_basic,
    "a = [1, 2, 3, 4]; a.keep_if {|x| x % 2 == 0}; puts a.join('-')",
    "2-4"
);
ruby_test!(
    test_keep_if_returns_self,
    "a = [1, 2]; puts a.keep_if {|x| true}.object_id == a.object_id",
    "true"
);
ruby_test!(
    test_keep_if_no_block,
    "puts [1, 2].keep_if.is_a?(Enumerator)",
    "true"
);
ruby_test!(
    test_keep_if_all,
    "a = [1, 2]; a.keep_if {|x| true}; puts a.length",
    "2"
);
ruby_test!(
    test_keep_if_none,
    "a = [1, 2]; a.keep_if {|x| false}; puts a.length",
    "0"
);
ruby_test!(
    test_delete_if_frozen,
    "# frozen_string_literal: true\na = [1].freeze; begin; a.delete_if {|x| true}; rescue FrozenError; puts 'err'; end",
    "err"
);
