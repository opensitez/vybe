use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_compact_basic, "puts [1, nil, 2, nil, 3].compact.join('-')", "1-2-3");
ruby_test!(test_compact_no_nils, "puts [1, 2, 3].compact.join('-')", "1-2-3");
ruby_test!(test_compact_all_nils, "puts [nil, nil].compact.length", "0");
ruby_test!(test_compact_empty, "puts [].compact.length", "0");
ruby_test!(test_compact_bang_mutates, "a = [1, nil, 2]; a.compact!; puts a.join('-')", "1-2");
ruby_test!(test_compact_bang_returns_nil, "a = [1, 2]; puts a.compact!.nil?", "true");
ruby_test!(test_compact_bang_all_nils, "a = [nil]; a.compact!; puts a.length", "0");
ruby_test!(test_compact_bang_returns_self, "a = [1, nil]; puts a.compact!.object_id == a.object_id", "true");
ruby_test!(test_compact_nested_nils, "puts [1, [nil, 2], nil].compact.inspect", "[1, [nil, 2]]"); // compact only removes top-level nils
ruby_test!(test_compact_with_false, "puts [1, false, nil, 2].compact.inspect", "[1, false, 2]"); // false is not removed
ruby_test!(test_compact_with_empty_string, "puts [1, '', nil].compact.inspect", "[1, \"\"]"); // empty string is not removed
