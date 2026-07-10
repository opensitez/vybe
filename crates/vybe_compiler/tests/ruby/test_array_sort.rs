use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_array_sort_basic, "puts [3, 1, 2].sort.join('-')", "1-2-3");
ruby_test!(test_array_sort_strings, "puts ['c', 'a', 'b'].sort.join('-')", "a-b-c");
ruby_test!(test_array_sort_block, "puts [3, 1, 2].sort { |a, b| b <=> a }.join('-')", "3-2-1");
ruby_test!(test_array_sort_bang, "a = [3, 1, 2]; a.sort!; puts a.join('-')", "1-2-3");
ruby_test!(test_array_sort_bang_block, "a = [3, 1, 2]; a.sort! { |x, y| y <=> x }; puts a.join('-')", "3-2-1");
ruby_test!(test_array_sort_by_basic, "puts ['apple', 'pear', 'fig'].sort_by { |word| word.length }.join('-')", "fig-pear-apple");
ruby_test!(test_array_sort_by_bang, "a = ['apple', 'pear', 'fig']; a.sort_by! { |word| word.length }; puts a.join('-')", "fig-pear-apple");
ruby_test!(test_array_sort_by_multiple, "puts ['apple', 'pear', 'fig', 'peach'].sort_by { |word| [word.length, word] }.join('-')", "fig-pear-apple-peach");
ruby_test!(test_array_sort_mixed_types, "begin; [1, 'a'].sort; rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_array_sort_by_enumerator, "puts ['apple', 'pear', 'fig'].sort_by.class.name", "Enumerator");
