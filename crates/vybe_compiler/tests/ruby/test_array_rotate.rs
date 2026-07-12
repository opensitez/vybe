macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_rotate_basic,
    "puts [1, 2, 3].rotate.join('-')",
    "2-3-1"
);
ruby_test!(
    test_rotate_count_positive,
    "puts [1, 2, 3].rotate(2).join('-')",
    "3-1-2"
);
ruby_test!(
    test_rotate_count_negative,
    "puts [1, 2, 3].rotate(-1).join('-')",
    "3-1-2"
);
ruby_test!(
    test_rotate_count_zero,
    "puts [1, 2, 3].rotate(0).join('-')",
    "1-2-3"
);
ruby_test!(
    test_rotate_count_larger_than_length,
    "puts [1, 2, 3].rotate(4).join('-')",
    "2-3-1"
);
ruby_test!(
    test_rotate_count_negative_larger,
    "puts [1, 2, 3].rotate(-4).join('-')",
    "3-1-2"
);
ruby_test!(test_rotate_empty, "puts [].rotate.length", "0");
ruby_test!(
    test_rotate_bang_mutates,
    "a = [1, 2, 3]; a.rotate!; puts a.join('-')",
    "2-3-1"
);
ruby_test!(
    test_rotate_bang_returns_self,
    "a = [1]; puts a.rotate!.object_id == a.object_id",
    "true"
);
ruby_test!(
    test_rotate_bang_count,
    "a = [1, 2, 3]; a.rotate!(-1); puts a.join('-')",
    "3-1-2"
);
ruby_test!(test_rotate_single_element, "puts [1].rotate.join('-')", "1");
ruby_test!(
    test_rotate_nil_element,
    "puts [nil, 2].rotate.inspect",
    "[2, nil]"
);
