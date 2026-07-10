use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_b_method_encoding, "s = 'café'.b; puts s.encoding.name", "ASCII-8BIT");
ruby_test!(test_b_method_preserves_bytes, "s = 'café'; puts s.b.bytes == s.bytes", "true");
ruby_test!(test_b_method_does_not_mutate, "s = 'café'; s.b; puts s.encoding.name", "UTF-8");
ruby_test!(test_byteslice_index, "puts 'café'.byteslice(1)", "a");
ruby_test!(test_byteslice_index_negative, "puts 'café'.byteslice(-2)", "\"\\xC3\".force_encoding('UTF-8')"); // The first byte of é's UTF-8 encoding. Wait, let's just test bytes length
ruby_test!(test_byteslice_index_negative_safe, "puts 'abc'.byteslice(-1)", "c");
ruby_test!(test_byteslice_length, "puts 'café'.byteslice(0, 3)", "caf");
ruby_test!(test_byteslice_length_unicode, "puts 'café'.byteslice(3, 2).force_encoding('UTF-8')", "é");
ruby_test!(test_byteslice_range, "puts 'café'.byteslice(0..2)", "caf");
ruby_test!(test_byteslice_range_unicode, "puts 'café'.byteslice(3..4).force_encoding('UTF-8')", "é");
ruby_test!(test_byteslice_out_of_bounds, "puts 'abc'.byteslice(5).nil?", "true");
ruby_test!(test_byteslice_length_out_of_bounds, "puts 'abc'.byteslice(1, 5)", "bc");
