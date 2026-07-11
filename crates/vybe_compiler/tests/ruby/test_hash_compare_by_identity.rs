
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_compare_by_identity_basic, "h = {}.compare_by_identity; h['a'] = 1; h['a'] = 2; puts h.length", "2"); // 'a' strings are different objects
ruby_test!(test_compare_by_identity_same_object, "h = {}.compare_by_identity; s = 'a'; h[s] = 1; h[s] = 2; puts h.length", "1");
ruby_test!(test_compare_by_identity_symbols, "h = {}.compare_by_identity; h[:a] = 1; h[:a] = 2; puts h.length", "1"); // symbols are same object
ruby_test!(test_compare_by_identity_integers, "h = {}.compare_by_identity; h[1] = 1; h[1] = 2; puts h.length", "1"); // integers are same object
ruby_test!(test_compare_by_identity_predicate, "puts {}.compare_by_identity?.to_s", "false");
ruby_test!(test_compare_by_identity_predicate_true, "puts {}.compare_by_identity.compare_by_identity?", "true");
ruby_test!(test_compare_by_identity_returns_self, "h = {}; puts h.compare_by_identity.object_id == h.object_id", "true");
ruby_test!(test_compare_by_identity_preserves_elements, "h = {'a' => 1}; h.compare_by_identity; puts h.length", "1");
ruby_test!(test_compare_by_identity_fetch, "h = {}.compare_by_identity; s = 'a'; h[s] = 1; puts h.fetch(s)", "1");
ruby_test!(test_compare_by_identity_fetch_different_object, "h = {}.compare_by_identity; h['a'] = 1; begin; h.fetch('a'); rescue KeyError; puts 'err'; end", "err");
