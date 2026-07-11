
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_string_modification_insert, "puts 'hello'.insert(2, 'x')", "hexllo");
ruby_test!(test_string_modification_insert_negative, "puts 'hello'.insert(-2, 'x')", "hellxo");
ruby_test!(test_string_modification_reverse, "puts 'hello'.reverse", "olleh");
ruby_test!(test_string_modification_reverse_bang, "s = 'hello'; s.reverse!; puts s", "olleh");
ruby_test!(test_string_modification_squeeze, "puts 'yellow moon'.squeeze", "yelow mon");
ruby_test!(test_string_modification_squeeze_char, "puts 'yellow moon'.squeeze('o')", "yellow mon");
ruby_test!(test_string_modification_squeeze_bang, "s = 'yellow moon'; s.squeeze!; puts s", "yelow mon");
ruby_test!(test_string_modification_tr, "puts 'hello'.tr('el', 'ip')", "hippo");
ruby_test!(test_string_modification_tr_bang, "s = 'hello'; s.tr!('el', 'ip'); puts s", "hippo");
ruby_test!(test_string_modification_tr_s, "puts 'hello'.tr_s('l', 'r')", "hero");
ruby_test!(test_string_modification_clear, "s = 'hello'; s.clear; puts s.length", "0");
ruby_test!(test_string_modification_replace, "s = 'hello'; s.replace('world'); puts s", "world");
ruby_test!(test_string_modification_concat, "s = 'hello'; s.concat(' world'); puts s", "hello world");
ruby_test!(test_string_modification_prepend, "s = 'hello'; s.prepend('say '); puts s", "say hello");
