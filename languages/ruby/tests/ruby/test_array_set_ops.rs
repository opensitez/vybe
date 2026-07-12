macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_array_intersection,
    "puts ([1, 2, 3] & [2, 3, 4]).join('-')",
    "2-3"
);
ruby_test!(
    test_array_union,
    "puts ([1, 2, 3] | [2, 3, 4]).join('-')",
    "1-2-3-4"
);
ruby_test!(
    test_array_difference,
    "puts ([1, 2, 3] - [2, 3, 4]).join('-')",
    "1"
);
ruby_test!(
    test_array_intersection_multiple,
    "puts [1, 2, 3].intersection([2, 3, 4], [3, 4, 5]).join('-')",
    "3"
);
ruby_test!(
    test_array_union_multiple,
    "puts [1, 2].union([2, 3], [3, 4]).join('-')",
    "1-2-3-4"
);
ruby_test!(
    test_array_difference_multiple,
    "puts [1, 2, 3, 4].difference([2], [4]).join('-')",
    "1-3"
);
ruby_test!(
    test_array_intersection_empty,
    "puts ([1, 2] & []).length",
    "0"
);
ruby_test!(
    test_array_union_empty,
    "puts ([1, 2] | []).join('-')",
    "1-2"
);
ruby_test!(
    test_array_difference_empty,
    "puts ([1, 2] - []).join('-')",
    "1-2"
);
ruby_test!(
    test_array_intersection_duplicates,
    "puts ([1, 1, 2] & [1, 2, 2]).join('-')",
    "1-2"
);
ruby_test!(
    test_array_union_duplicates,
    "puts ([1, 1, 2] | [1, 2, 2]).join('-')",
    "1-2"
);
ruby_test!(
    test_array_difference_duplicates,
    "puts ([1, 1, 2, 2, 3] - [2]).join('-')",
    "1-1-3"
); // difference removes all occurrences of 2
