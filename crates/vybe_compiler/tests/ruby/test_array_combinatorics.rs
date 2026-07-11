
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_array_combination_basic, "puts [1, 2, 3].combination(2).map { |a| a.join('') }.join('-')", "12-13-23");
ruby_test!(test_array_combination_zero, "puts [1, 2].combination(0).map { |a| a.join('') }.join('-')", "");
ruby_test!(test_array_combination_one, "puts [1, 2].combination(1).map { |a| a.join('') }.join('-')", "1-2");
ruby_test!(test_array_combination_all, "puts [1, 2].combination(2).map { |a| a.join('') }.join('-')", "12");
ruby_test!(test_array_combination_large, "puts [1, 2].combination(3).to_a.length", "0");
ruby_test!(test_array_permutation_basic, "puts [1, 2].permutation(2).map { |a| a.join('') }.join('-')", "12-21");
ruby_test!(test_array_permutation_zero, "puts [1, 2].permutation(0).map { |a| a.join('') }.join('-')", "");
ruby_test!(test_array_permutation_one, "puts [1, 2].permutation(1).map { |a| a.join('') }.join('-')", "1-2");
ruby_test!(test_array_permutation_large, "puts [1, 2].permutation(3).to_a.length", "0");
ruby_test!(test_array_permutation_all, "puts [1, 2, 3].permutation.to_a.length", "6");
