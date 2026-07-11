
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_repeated_permutation_basic, "puts [1, 2].repeated_permutation(2).map{|x| x.join('')}.join('-')", "11-12-21-22");
ruby_test!(test_repeated_permutation_one, "puts [1, 2].repeated_permutation(1).map{|x| x.join('')}.join('-')", "1-2");
ruby_test!(test_repeated_permutation_three, "puts [1, 2].repeated_permutation(3).map{|x| x.join('')}.join('-')", "111-112-121-122-211-212-221-222");
ruby_test!(test_repeated_permutation_zero, "puts [1, 2].repeated_permutation(0).to_a.inspect", "[[]]");
ruby_test!(test_repeated_permutation_negative, "puts [1, 2].repeated_permutation(-1).to_a.length", "0");
ruby_test!(test_repeated_permutation_empty_array, "puts [].repeated_permutation(1).to_a.length", "0");
ruby_test!(test_repeated_permutation_empty_array_zero, "puts [].repeated_permutation(0).to_a.inspect", "[[]]");
ruby_test!(test_repeated_permutation_with_block, "acc = []; [1].repeated_permutation(2) {|x| acc << x.join('')}; puts acc.join('-')", "11");
ruby_test!(test_repeated_permutation_returns_enumerator, "puts [1].repeated_permutation(1).is_a?(Enumerator)", "true");
