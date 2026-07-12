macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_merge_basic,
    "puts ({a: 1}.merge({b: 2}).keys.map(&:to_s).join('-'))",
    "a-b"
);
ruby_test!(
    test_merge_overwrites,
    "puts ({a: 1}.merge({a: 2})[:a])",
    "2"
);
ruby_test!(
    test_merge_multiple,
    "puts ({a: 1}.merge({b: 2}, {c: 3}).keys.map(&:to_s).join('-'))",
    "a-b-c"
); // ruby 2.6+
ruby_test!(
    test_merge_multiple_overwrites,
    "puts ({a: 1}.merge({a: 2}, {a: 3})[:a])",
    "3"
);
ruby_test!(
    test_merge_no_args,
    "puts ({a: 1}.merge.keys.map(&:to_s).join('-'))",
    "a"
);
ruby_test!(
    test_merge_empty,
    "puts ({a: 1}.merge({}).keys.map(&:to_s).join('-'))",
    "a"
);
ruby_test!(
    test_merge_with_block,
    "puts ({a: 1, b: 2}.merge({a: 3, c: 4}) {|k, v1, v2| v1 + v2}[:a])",
    "4"
);
ruby_test!(
    test_merge_block_not_called_for_new_keys,
    "acc = []; {a: 1}.merge({b: 2}) {|k, v1, v2| acc << k}; puts acc.length",
    "0"
);
ruby_test!(
    test_merge_bang_mutates,
    "h = {a: 1}; h.merge!({b: 2}); puts h.keys.map(&:to_s).join('-')",
    "a-b"
);
ruby_test!(
    test_merge_bang_returns_self,
    "h = {a: 1}; puts h.merge!({b: 2}).object_id == h.object_id",
    "true"
);
ruby_test!(
    test_merge_preserves_order,
    "puts ({a: 1, c: 3}.merge({b: 2}).keys.map(&:to_s).join('-'))",
    "a-c-b"
);
