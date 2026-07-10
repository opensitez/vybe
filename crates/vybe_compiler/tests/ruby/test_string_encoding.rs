use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_string_encoding_basic, "puts 'hello'.encoding.name", "UTF-8");
ruby_test!(test_string_encoding_force_encoding, "puts 'hello'.force_encoding('ASCII-8BIT').encoding.name", "ASCII-8BIT");
ruby_test!(test_string_encoding_encode, "puts 'hello'.encode('US-ASCII').encoding.name", "US-ASCII");
ruby_test!(test_string_encoding_valid_encoding, "puts 'hello'.valid_encoding?", "true");
ruby_test!(test_string_encoding_ascii_only, "puts 'hello'.ascii_only?", "true");
ruby_test!(test_string_encoding_unicode_ascii_only, "puts 'café'.ascii_only?", "false");
ruby_test!(test_string_encoding_b, "puts 'hello'.b.encoding.name", "ASCII-8BIT");
ruby_test!(test_string_encoding_bytesize, "puts 'café'.bytesize", "5"); // e with acute is 2 bytes in utf8
ruby_test!(test_string_encoding_length, "puts 'café'.length", "4");
ruby_test!(test_string_encoding_chr, "puts 'café'.chr", "c");
