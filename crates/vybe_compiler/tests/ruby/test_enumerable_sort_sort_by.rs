use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_sort_basic, "puts [3, 1, 2].sort.join('-')", "1-2-3");
ruby_test!(test_sort_strings, "puts ['c', 'a', 'b'].sort.join('-')", "a-b-c");
ruby_test!(test_sort_block, "puts [3, 1, 2].sort {|a, b| b <=> a}.join('-')", "3-2-1");
ruby_test!(test_sort_by_basic, "puts ['apple', 'pear', 'fig'].sort_by {|x| x.length}.join('-')", "fig-pear-apple");
ruby_test!(test_sort_by_no_block, "puts ['a'].sort_by.is_a?(Enumerator)", "true");
ruby_test!(test_sort_empty, "puts [].sort.length", "0");
ruby_test!(test_sort_by_empty, "puts [].sort_by {|x| x}.length", "0");
ruby_test!(test_sort_incomparable, "begin; [1, 'a'].sort; rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_sort_by_incomparable, "begin; ['a', 'b'].sort_by {|x| x == 'a' ? 1 : 'x'}; rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_sort_preserves_duplicates, "puts [2, 1, 2].sort.join('-')", "1-2-2");
ruby_test!(test_sort_hash, "puts ({b: 2, a: 1}.sort.map{|k, v| k.to_s}.join('-'))", "a-b"); // returns array of arrays sorted by key usually
