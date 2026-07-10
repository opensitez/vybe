use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_transcode_utf8_to_ascii, "s = 'abc'.encode('US-ASCII'); puts s.encoding.name", "US-ASCII");
ruby_test!(test_transcode_utf8_to_win1252, "s = 'café'.encode('Windows-1252'); puts s.encoding.name", "Windows-1252");
ruby_test!(test_transcode_win1252_to_utf8, "s = \"caf\\xE9\".force_encoding('Windows-1252').encode('UTF-8'); puts s.bytesize", "5"); // e is 2 bytes in UTF-8
ruby_test!(test_transcode_invalid_replace, "s = \"a\\xFFb\".force_encoding('UTF-8'); puts s.encode('UTF-8', invalid: :replace, replace: '*').bytes.join(',')", "97,42,98");
ruby_test!(test_transcode_undef_replace, "s = '😀'.encode('US-ASCII', undef: :replace, replace: '?'); puts s", "?");
ruby_test!(test_transcode_bang, "s = 'café'; s.encode!('Windows-1252'); puts s.encoding.name", "Windows-1252");
ruby_test!(test_transcode_same, "s = 'abc'; puts s.encode('UTF-8').object_id == s.object_id", "false"); // encode always returns new string
ruby_test!(test_transcode_same_bang, "s = 'abc'; id = s.object_id; s.encode!('UTF-8'); puts s.object_id == id", "true");
ruby_test!(test_transcode_to_utf16le, "s = 'a'.encode('UTF-16LE'); puts s.bytes.join(',')", "97,0");
ruby_test!(test_transcode_to_utf16be, "s = 'a'.encode('UTF-16BE'); puts s.bytes.join(',')", "0,97");
ruby_test!(test_transcode_from_utf16le, "s = \"a\\x00b\\x00\".force_encoding('UTF-16LE').encode('UTF-8'); puts s", "ab");
ruby_test!(test_transcode_invalid_ignore, "s = \"a\\xFFb\".force_encoding('UTF-8').encode('UTF-8', invalid: :replace, replace: ''); puts s", "ab");
