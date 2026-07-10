use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_array_sorting_sort, "puts [3, 1, 2].sort.join('-')", "1-2-3");
ruby_test!(test_array_sorting_sort_bang, "a = [3, 1, 2]; a.sort!; puts a.join('-')", "1-2-3");
ruby_test!(test_array_sorting_sort_block, "puts [3, 1, 2].sort { |a, b| b <=> a }.join('-')", "3-2-1");
ruby_test!(test_array_sorting_sort_bang_block, "a = [3, 1, 2]; a.sort! { |a, b| b <=> a }; puts a.join('-')", "3-2-1");
ruby_test!(test_array_sorting_sort_by, "puts %w[apple fig banana].sort_by { |word| word.length }.join('-')", "fig-apple-banana");
ruby_test!(test_array_sorting_sort_by_bang, "a = %w[apple fig banana]; a.sort_by! { |word| word.length }; puts a.join('-')", "fig-apple-banana");
ruby_test!(test_array_sorting_reverse, "puts [1, 2, 3].reverse.join('-')", "3-2-1");
ruby_test!(test_array_sorting_reverse_bang, "a = [1, 2, 3]; a.reverse!; puts a.join('-')", "3-2-1");
ruby_test!(test_array_sorting_shuffle, "a = [1, 2, 3]; puts a.shuffle.sort.join('-')", "1-2-3");
ruby_test!(test_array_sorting_shuffle_bang, "a = [1, 2, 3]; a.shuffle!; puts a.sort.join('-')", "1-2-3");
ruby_test!(test_array_sorting_shuffle_random, "r = Random.new(42); puts [1, 2, 3].shuffle(random: r).sort.join('-')", "1-2-3");
