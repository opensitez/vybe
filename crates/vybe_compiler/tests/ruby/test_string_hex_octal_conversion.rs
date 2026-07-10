use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_hex_to_i, "puts 'ff'.hex", "255");
ruby_test!(test_hex_to_i_prefix, "puts '0xff'.hex", "255");
ruby_test!(test_hex_to_i_invalid, "puts 'zff'.hex", "0");
ruby_test!(test_hex_to_i_partial, "puts 'ffz'.hex", "255");
ruby_test!(test_hex_to_i_negative, "puts '-ff'.hex", "-255");
ruby_test!(test_oct_to_i, "puts '10'.oct", "8");
ruby_test!(test_oct_to_i_prefix, "puts '010'.oct", "8");
ruby_test!(test_oct_to_i_invalid, "puts '8'.oct", "0"); // 8 is invalid octal
ruby_test!(test_oct_to_i_partial, "puts '108'.oct", "8");
ruby_test!(test_oct_to_i_negative, "puts '-10'.oct", "-8");
ruby_test!(test_oct_handles_hex_prefix, "puts '0xff'.oct", "255"); // oct handles hex prefix if present
ruby_test!(test_oct_handles_bin_prefix, "puts '0b10'.oct", "2"); // oct handles bin prefix if present
