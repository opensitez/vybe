macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_assoc_basic,
    "puts ({a: 1, b: 2}.assoc(:a).join('-'))",
    "a-1"
);
ruby_test!(test_assoc_missing, "puts ({a: 1}.assoc(:b).nil?)", "true");
ruby_test!(
    test_assoc_array_key,
    "puts ({[1, 2] => 3}.assoc([1, 2]).join('-'))",
    "1-2-3"
);
ruby_test!(
    test_assoc_string_key,
    "puts ({'a' => 1}.assoc('a').join('-'))",
    "a-1"
);
ruby_test!(
    test_rassoc_basic,
    "puts ({a: 1, b: 2}.rassoc(1).join('-'))",
    "a-1"
);
ruby_test!(test_rassoc_missing, "puts ({a: 1}.rassoc(2).nil?)", "true");
ruby_test!(
    test_rassoc_array_value,
    "puts ({a: [1, 2]}.rassoc([1, 2]).inspect)",
    "[:a, [1, 2]]"
);
ruby_test!(
    test_rassoc_string_value,
    "puts ({a: 'b'}.rassoc('b').join('-'))",
    "a-b"
);
ruby_test!(test_assoc_empty, "puts ({}).assoc(:a).nil?", "true");
ruby_test!(test_rassoc_empty, "puts ({}).rassoc(1).nil?", "true");
ruby_test!(
    test_assoc_nil_key,
    "puts ({nil => 1}.assoc(nil).inspect)",
    "[nil, 1]"
);
ruby_test!(
    test_rassoc_nil_value,
    "puts ({a: nil}.rassoc(nil).inspect)",
    "[:a, nil]"
);
