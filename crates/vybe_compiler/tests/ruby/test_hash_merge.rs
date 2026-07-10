use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_hash_merge_basic, "puts {a: 1}.merge({b: 2}).keys.sort.join('-')", "a-b");
ruby_test!(test_hash_merge_overlap, "puts {a: 1}.merge({a: 2})[:a]", "2");
ruby_test!(test_hash_merge_bang, "h = {a: 1}; h.merge!({b: 2}); puts h.keys.sort.join('-')", "a-b");
ruby_test!(test_hash_merge_block, "puts {a: 1}.merge({a: 2}) { |k, o, n| o + n }[:a]", "3");
ruby_test!(test_hash_merge_bang_block, "h = {a: 1}; h.merge!({a: 2}) { |k, o, n| o + n }; puts h[:a]", "3");
ruby_test!(test_hash_merge_multiple, "puts {a: 1}.merge({b: 2}, {c: 3}).keys.sort.join('-')", "a-b-c");
ruby_test!(test_hash_update, "h = {a: 1}; h.update({b: 2}); puts h.keys.sort.join('-')", "a-b");
