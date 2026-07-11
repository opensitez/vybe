
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_hash_compact, "puts {a: 1, b: nil}.compact.keys.join('-')", "a");
ruby_test!(test_hash_compact_bang, "h = {a: 1, b: nil}; h.compact!; puts h.keys.join('-')", "a");
ruby_test!(test_hash_compact_bang_nil, "puts {a: 1}.compact!.nil?", "true");
ruby_test!(test_hash_invert, "puts {a: 1, b: 2}.invert[1]", "a");
ruby_test!(test_hash_invert_duplicates, "puts {a: 1, b: 1}.invert[1] == :b || {a: 1, b: 1}.invert[1] == :a", "true");
ruby_test!(test_hash_invert_empty, "puts {}.invert.empty?", "true");
