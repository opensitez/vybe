
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_permutation_basic, "puts [1, 2].permutation.map{|x| x.join('')}.join('-')", "12-21");
ruby_test!(test_permutation_with_length, "puts [1, 2, 3].permutation(2).map{|x| x.join('')}.join('-')", "12-13-21-23-31-32");
ruby_test!(test_permutation_one, "puts [1, 2].permutation(1).map{|x| x.join('')}.join('-')", "1-2");
ruby_test!(test_permutation_all, "puts [1, 2, 3].permutation(3).map{|x| x.join('')}.join('-')", "123-132-213-231-312-321");
ruby_test!(test_permutation_zero, "puts [1, 2].permutation(0).to_a.inspect", "[[]]");
ruby_test!(test_permutation_out_of_bounds, "puts [1, 2].permutation(3).to_a.length", "0");
ruby_test!(test_permutation_negative, "puts [1, 2].permutation(-1).to_a.length", "0");
ruby_test!(test_permutation_empty_array, "puts [].permutation.to_a.inspect", "[[]]");
ruby_test!(test_permutation_empty_array_length, "puts [].permutation(1).to_a.length", "0");
ruby_test!(test_permutation_with_block, "acc = []; [1, 2].permutation {|x| acc << x.join('')}; puts acc.join('-')", "12-21");
ruby_test!(test_permutation_returns_enumerator, "puts [1].permutation.is_a?(Enumerator)", "true");
