macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_tr_basic, "puts 'hello'.tr('el', 'ip')", "hippo");
ruby_test!(
    test_tr_short_replacement,
    "puts 'hello'.tr('el', 'i')",
    "hiiio"
); // 'l' maps to 'i' (last char of replacement)
ruby_test!(test_tr_range, "puts 'hello'.tr('a-y', 'b-z')", "ifmmp");
ruby_test!(test_tr_negation, "puts 'hello'.tr('^aeiou', '*')", "*e**o");
ruby_test!(
    test_tr_negation_range,
    "puts 'hello 123'.tr('^a-z', ' ')",
    "hello    "
);
ruby_test!(
    test_tr_bang_mutates,
    "s = 'hello'; s.tr!('e', 'a'); puts s",
    "hallo"
);
ruby_test!(
    test_tr_bang_returns_nil_if_no_change,
    "s = 'hello'; puts s.tr!('z', 'x').nil?",
    "true"
);
ruby_test!(test_tr_s_basic, "puts 'hello'.tr_s('l', 'r')", "hero"); // translates 'll' to 'rr', then squeezes to 'r'
ruby_test!(
    test_tr_s_short_replacement,
    "puts 'hello'.tr_s('el', '*')",
    "h*o"
);
ruby_test!(
    test_tr_s_range,
    "puts 'aabbaabb'.tr_s('a-b', '1-2')",
    "1212"
);
ruby_test!(
    test_tr_s_negation,
    "puts 'hello   world'.tr_s('^a-z', ' ')",
    "hello world"
);
ruby_test!(
    test_tr_s_bang,
    "s = 'hello'; s.tr_s!('l', 'r'); puts s",
    "hero"
);
