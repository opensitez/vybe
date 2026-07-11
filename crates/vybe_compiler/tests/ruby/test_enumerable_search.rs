
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_enumerable_count, "puts [1, 2, 3, 2].count", "4");
ruby_test!(test_enumerable_count_item, "puts [1, 2, 3, 2].count(2)", "2");
ruby_test!(test_enumerable_count_block, "puts [1, 2, 3, 4].count { |x| x.even? }", "2");
ruby_test!(test_enumerable_find, "puts [1, 2, 3, 4].find { |x| x.even? }", "2");
ruby_test!(test_enumerable_find_ifnone, "puts [1, 3].find(-> { 'none' }) { |x| x.even? }", "none");
ruby_test!(test_enumerable_find_index, "puts [1, 2, 3, 4].find_index { |x| x.even? }", "1");
ruby_test!(test_enumerable_find_index_value, "puts [1, 2, 3].find_index(2)", "1");
ruby_test!(test_enumerable_first, "puts [1, 2, 3].first", "1");
ruby_test!(test_enumerable_first_n, "puts [1, 2, 3].first(2).join('-')", "1-2");
ruby_test!(test_enumerable_include, "puts [1, 2, 3].include?(2)", "true");
ruby_test!(test_enumerable_member, "puts [1, 2, 3].member?(2)", "true");
ruby_test!(test_enumerable_max, "puts [1, 3, 2].max", "3");
ruby_test!(test_enumerable_max_block, "puts ['a', 'ccc', 'bb'].max { |a, b| a.length <=> b.length }", "ccc");
ruby_test!(test_enumerable_max_n, "puts [1, 3, 2].max(2).join('-')", "3-2");
ruby_test!(test_enumerable_min, "puts [1, 3, 2].min", "1");
ruby_test!(test_enumerable_min_block, "puts ['a', 'ccc', 'bb'].min { |a, b| a.length <=> b.length }", "a");
ruby_test!(test_enumerable_min_n, "puts [1, 3, 2].min(2).join('-')", "1-2");
ruby_test!(test_enumerable_minmax, "puts [1, 3, 2].minmax.join('-')", "1-3");
ruby_test!(test_enumerable_minmax_block, "puts ['a', 'ccc', 'bb'].minmax { |a, b| a.length <=> b.length }.join('-')", "a-ccc");
