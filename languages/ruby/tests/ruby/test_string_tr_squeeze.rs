macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_string_tr_basic, "puts 'hello'.tr('el', 'ip')", "hippo");
ruby_test!(
    test_string_tr_range,
    "puts 'hello'.tr('a-z', 'A-Z')",
    "HELLO"
);
ruby_test!(test_string_tr_delete, "puts 'hello'.tr('l', '')", "heo");
ruby_test!(
    test_string_tr_bang,
    "s = 'hello'; s.tr!('e', 'a'); puts s",
    "hallo"
);
ruby_test!(
    test_string_tr_s_basic,
    "puts 'hello'.tr_s('l', 'p')",
    "hepo"
);
ruby_test!(
    test_string_tr_s_bang,
    "s = 'hello'; s.tr_s!('l', 'p'); puts s",
    "hepo"
);
ruby_test!(
    test_string_squeeze_basic,
    "puts 'yellow moon'.squeeze",
    "yelow mon"
);
ruby_test!(
    test_string_squeeze_args,
    "puts 'yellow moon'.squeeze('o')",
    "yellow mon"
);
ruby_test!(
    test_string_squeeze_bang,
    "s = 'yellow moon'; s.squeeze!; puts s",
    "yelow mon"
);
