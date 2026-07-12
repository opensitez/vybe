macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_shift_multiple,
    "a = [1, 2, 3]; puts a.shift(2).join('-')",
    "1-2"
);
ruby_test!(
    test_shift_multiple_mutates,
    "a = [1, 2, 3]; a.shift(2); puts a.join('-')",
    "3"
);
ruby_test!(
    test_shift_more_than_length,
    "a = [1]; puts a.shift(3).join('-')",
    "1"
);
ruby_test!(test_shift_zero, "a = [1]; puts a.shift(0).length", "0");
ruby_test!(
    test_shift_empty_multiple,
    "a = []; puts a.shift(2).length",
    "0"
);
ruby_test!(
    test_unshift_multiple,
    "a = [3]; a.unshift(1, 2); puts a.join('-')",
    "1-2-3"
);
ruby_test!(
    test_unshift_returns_self,
    "a = [1]; puts a.unshift(2).object_id == a.object_id",
    "true"
);
ruby_test!(
    test_unshift_zero_args,
    "a = [1]; a.unshift(); puts a.join('-')",
    "1"
);
ruby_test!(
    test_unshift_to_empty,
    "a = []; a.unshift(1, 2); puts a.join('-')",
    "1-2"
);
ruby_test!(
    test_prepend_alias,
    "a = [2]; a.prepend(1); puts a.join('-')",
    "1-2"
);
ruby_test!(
    test_shift_negative_error,
    "a = [1]; begin; a.shift(-1); rescue ArgumentError; puts 'err'; end",
    "err"
);
