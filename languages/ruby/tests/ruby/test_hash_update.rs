macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_update_alias,
    "h = {a: 1}; h.update({b: 2}); puts h.keys.map(&:to_s).join('-')",
    "a-b"
);
ruby_test!(
    test_update_overwrites,
    "h = {a: 1}; h.update({a: 2}); puts h[:a]",
    "2"
);
ruby_test!(
    test_update_multiple,
    "h = {a: 1}; h.update({b: 2}, {c: 3}); puts h.keys.map(&:to_s).join('-')",
    "a-b-c"
);
ruby_test!(
    test_update_multiple_overwrites,
    "h = {a: 1}; h.update({a: 2}, {a: 3}); puts h[:a]",
    "3"
);
ruby_test!(
    test_update_no_args,
    "h = {a: 1}; h.update; puts h.keys.map(&:to_s).join('-')",
    "a"
);
ruby_test!(
    test_update_empty,
    "h = {a: 1}; h.update({}); puts h.keys.map(&:to_s).join('-')",
    "a"
);
ruby_test!(
    test_update_with_block,
    "h = {a: 1, b: 2}; h.update({a: 3, c: 4}) {|k, v1, v2| v1 + v2}; puts h[:a]",
    "4"
);
ruby_test!(
    test_update_block_not_called_for_new_keys,
    "h = {a: 1}; acc = []; h.update({b: 2}) {|k, v1, v2| acc << k}; puts acc.length",
    "0"
);
ruby_test!(
    test_update_returns_self,
    "h = {a: 1}; puts h.update({b: 2}).object_id == h.object_id",
    "true"
);
ruby_test!(
    test_update_preserves_order,
    "h = {a: 1, c: 3}; h.update({b: 2}); puts h.keys.map(&:to_s).join('-')",
    "a-c-b"
);
