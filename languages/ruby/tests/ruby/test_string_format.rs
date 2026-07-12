macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_string_format_percent, "puts '%s %d' % ['a', 1]", "a 1");
ruby_test!(
    test_string_format_percent_hash,
    "puts '%{a} %{b}' % {a: 'x', b: 'y'}",
    "x y"
);
ruby_test!(
    test_string_format_center,
    "puts 'a'.center(5, '-')",
    "--a--"
);
ruby_test!(test_string_format_ljust, "puts 'a'.ljust(3, '-')", "a--");
ruby_test!(test_string_format_rjust, "puts 'a'.rjust(3, '-')", "--a");
ruby_test!(test_string_format_strip, "puts ' a '.strip", "a");
ruby_test!(test_string_format_lstrip, "puts ' a '.lstrip", "a ");
ruby_test!(test_string_format_rstrip, "puts ' a '.rstrip", " a");
ruby_test!(
    test_string_format_strip_bang,
    "s = ' a '; s.strip!; puts s",
    "a"
);
ruby_test!(test_string_format_chop, "puts 'abc\\n'.chop", "abc");
ruby_test!(
    test_string_format_chop_bang,
    "s = 'abc\\n'; s.chop!; puts s",
    "abc"
);
ruby_test!(test_string_format_chomp, "puts \"abc\\r\\n\".chomp", "abc");
ruby_test!(
    test_string_format_chomp_string,
    "puts 'abc'.chomp('c')",
    "ab"
);
ruby_test!(
    test_string_format_chomp_bang,
    "s = 'abc\\n'; s.chomp!; puts s",
    "abc"
);
