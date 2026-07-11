
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_hash_compact_basic, "puts ({a: 1, b: nil, c: 3}.compact.keys.join('-'))", "a-c");
ruby_test!(test_hash_compact_bang_basic, "h = {a: 1, b: nil, c: 3}; h.compact!; puts h.keys.join('-')", "a-c");
ruby_test!(test_hash_compact_bang_no_change, "h = {a: 1}; puts h.compact!.nil?", "true");
ruby_test!(test_hash_compact_empty, "puts {}.compact.length", "0");
ruby_test!(test_hash_compact_all_nil, "puts ({a: nil, b: nil}.compact.length)", "0");
ruby_test!(test_hash_compact_false_remains, "puts ({a: false, b: nil}.compact.keys.join('-'))", "a");
