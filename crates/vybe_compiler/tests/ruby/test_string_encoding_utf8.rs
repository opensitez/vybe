use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_utf8_encoding_name, "s = 'abc'; puts s.encoding.name", "UTF-8");
ruby_test!(test_utf8_length, "s = 'café'; puts s.length", "4"); // 4 characters
ruby_test!(test_utf8_bytesize, "s = 'café'; puts s.bytesize", "5"); // e is 2 bytes
ruby_test!(test_utf8_chars, "s = 'café'; puts s.chars.join('-')", "c-a-f-é");
ruby_test!(test_utf8_valid_encoding, "s = 'café'; puts s.valid_encoding?", "true");
ruby_test!(test_utf8_invalid_bytes, "s = \"a\\xFFb\".force_encoding('UTF-8'); puts s.valid_encoding?", "false");
ruby_test!(test_utf8_slice_char, "s = 'café'; puts s[3]", "é");
ruby_test!(test_utf8_codepoints, "s = 'café'; puts s.codepoints.last", "233");
ruby_test!(test_utf8_scrub, "s = \"a\\xFFb\".force_encoding('UTF-8'); puts s.scrub('*')", "a*b");
ruby_test!(test_utf8_unicode_escape, "puts \"\\u{1F600}\".length", "1"); // 😀 emoji
ruby_test!(test_utf8_unicode_bytesize, "puts \"\\u{1F600}\".bytesize", "4");
ruby_test!(test_utf8_concat, "s1 = 'a'; s2 = 'é'; puts (s1+s2).encoding.name", "UTF-8");
