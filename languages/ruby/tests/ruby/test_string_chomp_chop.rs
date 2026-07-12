macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_string_chomp_basic, "puts 'hello\\n'.chomp", "hello");
ruby_test!(test_string_chomp_crlf, "puts 'hello\\r\\n'.chomp", "hello");
ruby_test!(test_string_chomp_string, "puts 'hello'.chomp('lo')", "hel");
ruby_test!(
    test_string_chomp_bang,
    "s = 'hello\\n'; s.chomp!; puts s",
    "hello"
);
ruby_test!(
    test_string_chomp_bang_no_change,
    "s = 'hello'; puts s.chomp!.nil?",
    "true"
);
ruby_test!(test_string_chop_basic, "puts 'hello'.chop", "hell");
ruby_test!(test_string_chop_crlf, "puts 'hello\\r\\n'.chop", "hello");
ruby_test!(
    test_string_chop_bang,
    "s = 'hello'; s.chop!; puts s",
    "hell"
);
ruby_test!(
    test_string_chop_bang_empty,
    "s = ''; puts s.chop!.nil?",
    "true"
);
