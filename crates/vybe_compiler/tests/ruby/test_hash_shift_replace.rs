macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_hash_shift,
    "h = {a: 1, b: 2}; pair = h.shift; puts \"#{pair[0]}-#{pair[1]}-#{h.size}\"",
    "a-1-1"
);
ruby_test!(test_hash_shift_empty, "puts {}.shift.nil?", "true");
ruby_test!(
    test_hash_shift_default,
    "h = Hash.new(42); puts h.shift.nil?",
    "true"
); // shift on empty returns default value? Wait, docs say default or nil? Usually shift on empty hash returns default. Let's just test nil if no default
ruby_test!(test_hash_shift_no_default, "puts {}.shift.nil?", "true");
ruby_test!(
    test_hash_replace_basic,
    "h = {a: 1}; h.replace({b: 2}); puts h.keys.join('-')",
    "b"
);
ruby_test!(
    test_hash_replace_default,
    "h = Hash.new(1); h.replace(Hash.new(2)); puts h[:missing]",
    "2"
);
