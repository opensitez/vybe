macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_rehash_basic,
    "h = {}; s = 'a'; h[s] = 1; s.upcase!; h.rehash; puts h[s]",
    "1"
);
ruby_test!(
    test_rehash_returns_self,
    "h = {}; puts h.rehash.object_id == h.object_id",
    "true"
);
ruby_test!(
    test_rehash_duplicate_keys,
    "h = {}; a = [1]; b = [2]; h[a] = 1; h[b] = 2; a[0] = 2; h.rehash; puts h.length",
    "2"
); // both still exist, but they have same hash/eql
ruby_test!(
    test_rehash_duplicate_keys_access,
    "h = {}; a = [1]; b = [2]; h[a] = 1; h[b] = 2; a[0] = 2; h.rehash; puts h[[2]]",
    "2"
); // usually returns the one added last, or first, behavior is defined but edge case
ruby_test!(test_rehash_empty, "puts {}.rehash.length", "0");
ruby_test!(
    test_rehash_frozen_error,
    "# frozen_string_literal: true\nh = {a: 1}.freeze; begin; h.rehash; rescue FrozenError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_rehash_frozen_string_key,
    "# frozen_string_literal: true\nh = {'a' => 1}; h.rehash; puts h['a']",
    "1"
);
