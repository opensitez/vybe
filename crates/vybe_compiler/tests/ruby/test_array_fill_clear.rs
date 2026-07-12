macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_array_fill_basic,
    "puts [1, 2, 3].fill('x').join('-')",
    "x-x-x"
);
ruby_test!(
    test_array_fill_range,
    "puts [1, 2, 3, 4].fill('x', 1..2).join('-')",
    "1-x-x-4"
);
ruby_test!(
    test_array_fill_start,
    "puts [1, 2, 3].fill('x', 1).join('-')",
    "1-x-x"
);
ruby_test!(
    test_array_fill_start_length,
    "puts [1, 2, 3, 4].fill('x', 1, 2).join('-')",
    "1-x-x-4"
);
ruby_test!(
    test_array_fill_block,
    "puts [1, 2, 3].fill { |i| i * 2 }.join('-')",
    "0-2-4"
);
ruby_test!(
    test_array_fill_block_start_length,
    "puts [1, 2, 3, 4].fill(1, 2) { |i| i * 2 }.join('-')",
    "1-2-4-4"
);
ruby_test!(
    test_array_clear,
    "a = [1, 2, 3]; a.clear; puts a.length",
    "0"
);
ruby_test!(
    test_array_replace,
    "a = [1, 2]; a.replace([3, 4, 5]); puts a.join('-')",
    "3-4-5"
);
ruby_test!(
    test_array_insert_basic,
    "a = [1, 2]; a.insert(1, 'x'); puts a.join('-')",
    "1-x-2"
);
ruby_test!(
    test_array_insert_multiple,
    "a = [1, 2]; a.insert(1, 'x', 'y'); puts a.join('-')",
    "1-x-y-2"
);
ruby_test!(
    test_array_insert_negative,
    "a = [1, 2, 3]; a.insert(-2, 'x'); puts a.join('-')",
    "1-2-x-3"
);
