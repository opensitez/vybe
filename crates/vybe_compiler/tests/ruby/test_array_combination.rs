macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_combination_basic,
    "puts [1, 2, 3].combination(2).map{|x| x.join('')}.join('-')",
    "12-13-23"
);
ruby_test!(
    test_combination_one,
    "puts [1, 2, 3].combination(1).map{|x| x.join('')}.join('-')",
    "1-2-3"
);
ruby_test!(
    test_combination_all,
    "puts [1, 2, 3].combination(3).map{|x| x.join('')}.join('-')",
    "123"
);
ruby_test!(
    test_combination_zero,
    "puts [1, 2, 3].combination(0).map{|x| x.join('')}.join('-')",
    ""
); // yields one empty array
ruby_test!(
    test_combination_out_of_bounds,
    "puts [1, 2, 3].combination(4).to_a.length",
    "0"
);
ruby_test!(
    test_combination_negative,
    "puts [1, 2, 3].combination(-1).to_a.length",
    "0"
); // negative length returns empty enumerator
ruby_test!(
    test_combination_empty_array,
    "puts [].combination(1).to_a.length",
    "0"
);
ruby_test!(
    test_combination_empty_array_zero,
    "puts [].combination(0).to_a.inspect",
    "[[]]"
);
ruby_test!(
    test_combination_with_block,
    "acc = []; [1, 2].combination(1) {|x| acc << x[0]}; puts acc.join('-')",
    "1-2"
);
ruby_test!(
    test_combination_returns_enumerator,
    "puts [1, 2].combination(1).is_a?(Enumerator)",
    "true"
);
ruby_test!(
    test_combination_no_block_returns_enumerator,
    "puts [1, 2].combination(1).to_a.inspect",
    "[[1], [2]]"
);
