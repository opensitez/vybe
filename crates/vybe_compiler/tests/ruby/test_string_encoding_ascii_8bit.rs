use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_ascii_8bit_encoding, "s = 'a'.force_encoding('ASCII-8BIT'); puts s.encoding.name", "ASCII-8BIT");
ruby_test!(test_ascii_8bit_concat, "s1 = 'a'.force_encoding('ASCII-8BIT'); s2 = 'b'.force_encoding('ASCII-8BIT'); puts (s1+s2).encoding.name", "ASCII-8BIT");
ruby_test!(test_ascii_8bit_bytes, "s = 'abc'.force_encoding('ASCII-8BIT'); puts s.bytes.join(',')", "97,98,99");
ruby_test!(test_ascii_8bit_valid, "s = \"\\xFF\".force_encoding('ASCII-8BIT'); puts s.valid_encoding?", "true");
ruby_test!(test_ascii_8bit_length, "s = \"\\xFF\\xFE\".force_encoding('ASCII-8BIT'); puts s.length", "2");
ruby_test!(test_ascii_8bit_slice, "s = 'abcdef'.force_encoding('ASCII-8BIT'); puts s[1..2]", "bc");
ruby_test!(test_ascii_8bit_eql, "s1 = 'a'.force_encoding('ASCII-8BIT'); s2 = 'a'.force_encoding('UTF-8'); puts s1 == s2", "true"); // 'a' is 7-bit, so == is true
ruby_test!(test_ascii_8bit_eql_high, "s1 = \"\\xFF\".force_encoding('ASCII-8BIT'); s2 = \"\\xFF\".force_encoding('UTF-8'); puts s1 == s2", "false");
ruby_test!(test_ascii_8bit_to_s, "s = 'x'.force_encoding('ASCII-8BIT'); puts s.to_s.encoding.name", "ASCII-8BIT");
ruby_test!(test_ascii_8bit_inspect, "s = \"\\xFF\".force_encoding('ASCII-8BIT'); puts s.inspect", "\"\\xFF\"");
ruby_test!(test_ascii_8bit_b_method, "s = 'abc'.b; puts s.encoding.name", "ASCII-8BIT");
ruby_test!(test_ascii_8bit_chr, "puts 97.chr(Encoding::ASCII_8BIT).encoding.name", "ASCII-8BIT");
