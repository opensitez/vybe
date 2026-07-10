use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_repeated_combination_basic, "puts [1, 2].repeated_combination(2).map{|x| x.join('')}.join('-')", "11-12-22");
ruby_test!(test_repeated_combination_one, "puts [1, 2].repeated_combination(1).map{|x| x.join('')}.join('-')", "1-2");
ruby_test!(test_repeated_combination_three, "puts [1, 2].repeated_combination(3).map{|x| x.join('')}.join('-')", "111-112-122-222"); // wait, combination order: 111, 112, 122, 222
ruby_test!(test_repeated_combination_zero, "puts [1, 2].repeated_combination(0).to_a.inspect", "[[]]");
ruby_test!(test_repeated_combination_negative, "puts [1, 2].repeated_combination(-1).to_a.length", "0");
ruby_test!(test_repeated_combination_empty_array, "puts [].repeated_combination(1).to_a.length", "0");
ruby_test!(test_repeated_combination_empty_array_zero, "puts [].repeated_combination(0).to_a.inspect", "[[]]");
ruby_test!(test_repeated_combination_with_block, "acc = []; [1].repeated_combination(2) {|x| acc << x.join('')}; puts acc.join('-')", "11");
ruby_test!(test_repeated_combination_returns_enumerator, "puts [1].repeated_combination(1).is_a?(Enumerator)", "true");
