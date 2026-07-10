use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_range_min, "puts (1..5).min", "1");
ruby_test!(test_range_min_block, "puts (1..5).min { |a, b| b <=> a }", "5");
ruby_test!(test_range_max, "puts (1..5).max", "5");
ruby_test!(test_range_max_exclusive, "puts (1...5).max", "4");
ruby_test!(test_range_max_block, "puts (1..5).max { |a, b| b <=> a }", "1");
ruby_test!(test_range_minmax, "puts (1..5).minmax.join('-')", "1-5");
ruby_test!(test_range_minmax_exclusive, "puts (1...5).minmax.join('-')", "1-4");
ruby_test!(test_range_first, "puts (1..5).first", "1");
ruby_test!(test_range_first_n, "puts (1..5).first(2).join('-')", "1-2");
ruby_test!(test_range_last, "puts (1..5).last", "5");
ruby_test!(test_range_last_exclusive, "puts (1...5).last", "5"); // last on exclusive still returns the end value
ruby_test!(test_range_last_n, "puts (1..5).last(2).join('-')", "4-5");
ruby_test!(test_range_last_n_exclusive, "puts (1...5).last(2).join('-')", "3-4");
ruby_test!(test_range_size_integer, "puts (1..5).size", "5");
ruby_test!(test_range_size_integer_exclusive, "puts (1...5).size", "4");
ruby_test!(test_range_size_string, "puts ('a'..'z').size.nil?", "true");
ruby_test!(test_range_to_a, "puts (1..3).to_a.join('-')", "1-2-3");
ruby_test!(test_range_entries, "puts (1..3).entries.join('-')", "1-2-3");
