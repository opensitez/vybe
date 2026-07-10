use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_string_hex_basic, "puts '1a'.hex", "26");
ruby_test!(test_string_hex_prefix, "puts '0x1a'.hex", "26");
ruby_test!(test_string_hex_invalid, "puts 'z'.hex", "0");
ruby_test!(test_string_oct_basic, "puts '10'.oct", "8");
ruby_test!(test_string_oct_prefix, "puts '010'.oct", "8");
ruby_test!(test_string_oct_hex_prefix, "puts '0x10'.oct", "16");
ruby_test!(test_string_oct_bin_prefix, "puts '0b10'.oct", "2");
ruby_test!(test_string_oct_invalid, "puts '8'.oct", "0");
