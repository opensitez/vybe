use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_min_basic, "puts [3, 1, 2].min", "1");
ruby_test!(test_min_strings, "puts ['c', 'a', 'b'].min", "a");
ruby_test!(test_min_block, "puts [3, 1, 2].min {|a, b| b <=> a}", "3");
ruby_test!(test_min_count, "puts [3, 1, 2].min(2).join('-')", "1-2"); // lowest 2 elements sorted
ruby_test!(test_min_count_block, "puts [3, 1, 2].min(2) {|a, b| b <=> a}.join('-')", "3-2");
ruby_test!(test_min_empty, "puts [].min.nil?", "true");
ruby_test!(test_min_empty_count, "puts [].min(2).length", "0");
ruby_test!(test_max_basic, "puts [3, 1, 2].max", "3");
ruby_test!(test_max_strings, "puts ['c', 'a', 'b'].max", "c");
ruby_test!(test_max_block, "puts [3, 1, 2].max {|a, b| b <=> a}", "1");
ruby_test!(test_max_count, "puts [3, 1, 2].max(2).join('-')", "3-2"); // highest 2 elements sorted descending
ruby_test!(test_max_count_block, "puts [3, 1, 2].max(2) {|a, b| b <=> a}.join('-')", "1-2");
ruby_test!(test_max_empty, "puts [].max.nil?", "true");
ruby_test!(test_minmax_basic, "puts [3, 1, 2].minmax.join('-')", "1-3");
ruby_test!(test_minmax_block, "puts [3, 1, 2].minmax {|a, b| b <=> a}.join('-')", "3-1");
ruby_test!(test_minmax_empty, "puts [].minmax.inspect", "[nil, nil]");
