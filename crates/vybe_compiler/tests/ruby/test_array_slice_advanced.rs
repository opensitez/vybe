use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_array_slice_advanced_range, "puts [1, 2, 3, 4, 5].slice(1..3).join('-')", "2-3-4");
ruby_test!(test_array_slice_advanced_range_exclusive, "puts [1, 2, 3, 4, 5].slice(1...3).join('-')", "2-3");
ruby_test!(test_array_slice_advanced_range_negative, "puts [1, 2, 3, 4, 5].slice(-3..-1).join('-')", "3-4-5");
ruby_test!(test_array_slice_advanced_length, "puts [1, 2, 3, 4, 5].slice(1, 3).join('-')", "2-3-4");
ruby_test!(test_array_slice_advanced_length_too_large, "puts [1, 2, 3].slice(1, 10).join('-')", "2-3");
ruby_test!(test_array_slice_advanced_length_negative, "puts [1, 2, 3].slice(1, -1).nil?", "true");
ruby_test!(test_array_slice_advanced_index_out_of_bounds, "puts [1, 2].slice(5, 1).nil?", "true");
ruby_test!(test_array_slice_advanced_bang_range, "a = [1, 2, 3, 4]; puts a.slice!(1..2).join('-'); puts a.join('-')", "2-3\n1-4");
ruby_test!(test_array_slice_advanced_bang_length, "a = [1, 2, 3, 4]; puts a.slice!(1, 2).join('-'); puts a.join('-')", "2-3\n1-4");
ruby_test!(test_array_drop, "puts [1, 2, 3].drop(1).join('-')", "2-3");
ruby_test!(test_array_drop_large, "puts [1, 2].drop(5).join('-')", "");
ruby_test!(test_array_drop_while, "puts [1, 2, 3, 1].drop_while { |x| x < 3 }.join('-')", "3-1");
ruby_test!(test_array_take, "puts [1, 2, 3].take(2).join('-')", "1-2");
ruby_test!(test_array_take_large, "puts [1, 2].take(5).join('-')", "1-2");
ruby_test!(test_array_take_while, "puts [1, 2, 3, 1].take_while { |x| x < 3 }.join('-')", "1-2");
