macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_array_combination_1,
    "acc = []; [1, 2, 3].combination(1) { |c| acc << c.join }; puts acc.join('-')",
    "1-2-3"
);
ruby_test!(
    test_array_combination_2,
    "acc = []; [1, 2, 3].combination(2) { |c| acc << c.join }; puts acc.join('-')",
    "12-13-23"
);
ruby_test!(
    test_array_combination_3,
    "acc = []; [1, 2, 3].combination(3) { |c| acc << c.join }; puts acc.join('-')",
    "123"
);
ruby_test!(
    test_array_combination_4,
    "acc = []; [1, 2, 3].combination(4) { |c| acc << c.join }; puts acc.length",
    "0"
);
ruby_test!(
    test_array_combination_0,
    "acc = []; [1, 2, 3].combination(0) { |c| acc << c.join }; puts acc.join('-')",
    ""
); // yields empty array
ruby_test!(
    test_array_combination_enumerator,
    "puts [1, 2].combination(1).class.name",
    "Enumerator"
);
ruby_test!(
    test_array_permutation_2,
    "acc = []; [1, 2].permutation(2) { |p| acc << p.join }; puts acc.join('-')",
    "12-21"
);
ruby_test!(
    test_array_permutation_1,
    "acc = []; [1, 2].permutation(1) { |p| acc << p.join }; puts acc.join('-')",
    "1-2"
);
ruby_test!(
    test_array_permutation_0,
    "acc = []; [1, 2].permutation(0) { |p| acc << p.join }; puts acc.join('-')",
    ""
);
ruby_test!(
    test_array_permutation_no_arg,
    "acc = []; [1, 2].permutation { |p| acc << p.join }; puts acc.join('-')",
    "12-21"
);
ruby_test!(
    test_array_permutation_enumerator,
    "puts [1, 2].permutation.class.name",
    "Enumerator"
);
