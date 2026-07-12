macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_hash_manipulation_store,
    "h = {a: 1}; h.store(:b, 2); puts h.keys.join('-')",
    "a-b"
);
ruby_test!(
    test_hash_manipulation_delete,
    "h = {a: 1, b: 2}; h.delete(:a); puts h.keys.join('-')",
    "b"
);
ruby_test!(
    test_hash_manipulation_delete_missing,
    "h = {a: 1}; puts h.delete(:b).nil?",
    "true"
);
ruby_test!(
    test_hash_manipulation_delete_block,
    "h = {a: 1}; puts h.delete(:b) { |k| \"missing #{k}\" }",
    "missing b"
);
ruby_test!(
    test_hash_manipulation_delete_if,
    "h = {a: 1, b: 2}; h.delete_if { |k, v| v > 1 }; puts h.keys.join('-')",
    "a"
);
ruby_test!(
    test_hash_manipulation_keep_if,
    "h = {a: 1, b: 2}; h.keep_if { |k, v| v > 1 }; puts h.keys.join('-')",
    "b"
);
ruby_test!(
    test_hash_manipulation_reject_bang,
    "h = {a: 1, b: 2}; h.reject! { |k, v| v > 1 }; puts h.keys.join('-')",
    "a"
);
ruby_test!(
    test_hash_manipulation_select_bang,
    "h = {a: 1, b: 2}; h.select! { |k, v| v > 1 }; puts h.keys.join('-')",
    "b"
);
ruby_test!(
    test_hash_manipulation_clear,
    "h = {a: 1}; h.clear; puts h.empty?",
    "true"
);
ruby_test!(
    test_hash_manipulation_shift,
    "h = {a: 1, b: 2}; puts h.shift.join('-')",
    "a-1"
);
ruby_test!(
    test_hash_manipulation_shift_empty,
    "h = {}; puts h.shift.nil?",
    "true"
);
ruby_test!(
    test_hash_manipulation_invert,
    "h = {a: 1, b: 2}; puts h.invert[1]",
    "a"
);
