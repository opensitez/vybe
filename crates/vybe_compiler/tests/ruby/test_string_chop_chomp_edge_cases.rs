
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_chop_basic, "puts 'hello'.chop", "hell");
ruby_test!(test_chop_newline, "puts \"hello\\n\".chop", "hello");
ruby_test!(test_chop_crlf, "puts \"hello\\r\\n\".chop", "hello"); // chop removes both \r\n
ruby_test!(test_chop_empty, "puts ''.chop", "");
ruby_test!(test_chop_single, "puts 'a'.chop", "");
ruby_test!(test_chop_bang, "s = 'abc'; s.chop!; puts s", "ab");
ruby_test!(test_chomp_basic, "puts \"hello\\n\".chomp", "hello");
ruby_test!(test_chomp_crlf, "puts \"hello\\r\\n\".chomp", "hello");
ruby_test!(test_chomp_no_newline, "puts 'hello'.chomp", "hello");
ruby_test!(test_chomp_custom, "puts 'hello'.chomp('llo')", "he");
ruby_test!(test_chomp_empty, "puts ''.chomp", "");
ruby_test!(test_chomp_bang, "s = \"abc\\n\"; s.chomp!; puts s", "abc");
