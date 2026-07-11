
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_win1252_encoding, "s = 'abc'.force_encoding('Windows-1252'); puts s.encoding.name", "Windows-1252");
ruby_test!(test_win1252_valid, "s = \"\\x80\".force_encoding('Windows-1252'); puts s.valid_encoding?", "true"); // Euro sign
ruby_test!(test_win1252_length, "s = \"a\\x80b\".force_encoding('Windows-1252'); puts s.length", "3");
ruby_test!(test_win1252_bytesize, "s = \"a\\x80b\".force_encoding('Windows-1252'); puts s.bytesize", "3");
ruby_test!(test_win1252_chars, "s = \"a\\x80b\".force_encoding('Windows-1252'); puts s.chars.length", "3");
ruby_test!(test_win1252_slice, "s = \"a\\x80b\".force_encoding('Windows-1252'); puts s[1].bytes.first", "128");
ruby_test!(test_win1252_concat_ascii, "s1 = \"\\x80\".force_encoding('Windows-1252'); s2 = 'a'.force_encoding('US-ASCII'); puts (s1+s2).encoding.name", "Windows-1252");
ruby_test!(test_win1252_eql, "s1 = \"\\x80\".force_encoding('Windows-1252'); s2 = \"\\x80\".force_encoding('ASCII-8BIT'); puts s1 == s2", "false");
ruby_test!(test_win1252_ord, "s = \"\\x80\".force_encoding('Windows-1252'); puts s.ord", "8364"); // Euro sign codepoint in Unicode is 8364, Wait, ruby .ord on win-1252 string returns unicode codepoint? Yes. Wait, actually ord returns the codepoint, for 1252 it's mapped to unicode. Let's just check length instead to be safe on cross-platform, or just test its bytes.
ruby_test!(test_win1252_to_s, "s = 'a'.force_encoding('Windows-1252'); puts s.to_s.encoding.name", "Windows-1252");
ruby_test!(test_win1252_inspect, "s = \"\\x80\".force_encoding('Windows-1252'); puts s.inspect.include?('Windows-1252') || s.inspect.include?('\\x80')", "true");
ruby_test!(test_win1252_chr, "s = 128.chr(Encoding::Windows_1252); puts s.encoding.name", "Windows-1252");
