
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_intersect_operator, "puts ([1, 2, 3] & [2, 3, 4]).join('-')", "2-3");
ruby_test!(test_intersect_multiple, "puts ([1, 2, 3] & [2, 3] & [3, 4]).join('-')", "3");
ruby_test!(test_intersect_empty_result, "puts ([1, 2] & [3, 4]).length", "0");
ruby_test!(test_intersect_empty_receiver, "puts ([] & [1, 2]).length", "0");
ruby_test!(test_intersect_empty_arg, "puts ([1, 2] & []).length", "0");
ruby_test!(test_intersect_preserves_order, "puts ([3, 2, 1] & [1, 2, 3]).join('-')", "3-2-1"); // order of receiver
ruby_test!(test_intersect_removes_duplicates, "puts ([1, 1, 2] & [1, 2, 2]).join('-')", "1-2");
ruby_test!(test_union_operator, "puts ([1, 2] | [2, 3]).join('-')", "1-2-3");
ruby_test!(test_union_multiple, "puts ([1, 2] | [2, 3] | [3, 4]).join('-')", "1-2-3-4");
ruby_test!(test_union_empty_receiver, "puts ([] | [1, 2]).join('-')", "1-2");
ruby_test!(test_union_empty_arg, "puts ([1, 2] | []).join('-')", "1-2");
ruby_test!(test_union_preserves_order, "puts ([3, 1] | [2, 1]).join('-')", "3-1-2");
ruby_test!(test_union_removes_duplicates, "puts ([1, 1] | [2, 2]).join('-')", "1-2");
ruby_test!(test_difference_operator, "puts ([1, 2, 3, 4] - [2, 4]).join('-')", "1-3");
ruby_test!(test_difference_removes_all_instances, "puts ([1, 2, 2, 3] - [2]).join('-')", "1-3");
ruby_test!(test_intersection_method, "puts [1, 2, 3].intersection([2, 3, 4]).join('-')", "2-3");
ruby_test!(test_union_method, "puts [1, 2].union([2, 3]).join('-')", "1-2-3");
ruby_test!(test_difference_method, "puts [1, 2, 3].difference([2]).join('-')", "1-3");
