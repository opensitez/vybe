macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_enumerable_all,
    "puts [1, 2, 3].all? { |x| x > 0 }",
    "true"
);
ruby_test!(
    test_enumerable_all_false,
    "puts [1, 2, 3].all? { |x| x > 1 }",
    "false"
);
ruby_test!(
    test_enumerable_all_no_block,
    "puts [1, true, 'a'].all?",
    "true"
);
ruby_test!(
    test_enumerable_all_pattern,
    "puts [1, 2, 3].all?(Integer)",
    "true"
);
ruby_test!(
    test_enumerable_any,
    "puts [1, 2, 3].any? { |x| x > 2 }",
    "true"
);
ruby_test!(
    test_enumerable_any_false,
    "puts [1, 2, 3].any? { |x| x > 3 }",
    "false"
);
ruby_test!(
    test_enumerable_any_no_block,
    "puts [nil, false, 1].any?",
    "true"
);
ruby_test!(
    test_enumerable_any_pattern,
    "puts ['a', 'b'].any?(/b/)",
    "true"
);
ruby_test!(
    test_enumerable_none,
    "puts [1, 2, 3].none? { |x| x > 3 }",
    "true"
);
ruby_test!(
    test_enumerable_none_false,
    "puts [1, 2, 3].none? { |x| x > 2 }",
    "false"
);
ruby_test!(
    test_enumerable_none_no_block,
    "puts [nil, false].none?",
    "true"
);
ruby_test!(
    test_enumerable_none_pattern,
    "puts ['a', 'b'].none?(/c/)",
    "true"
);
ruby_test!(
    test_enumerable_one,
    "puts [1, 2, 3].one? { |x| x == 2 }",
    "true"
);
ruby_test!(
    test_enumerable_one_false_zero,
    "puts [1, 2, 3].one? { |x| x == 4 }",
    "false"
);
ruby_test!(
    test_enumerable_one_false_many,
    "puts [1, 2, 3].one? { |x| x > 0 }",
    "false"
);
ruby_test!(
    test_enumerable_one_no_block,
    "puts [nil, 1, false].one?",
    "true"
);
ruby_test!(
    test_enumerable_one_pattern,
    "puts ['a', 'b', 'c'].one?(/b/)",
    "true"
);
