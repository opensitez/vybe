
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_numeric_to_i, "puts 1.5.to_i", "1");
ruby_test!(test_numeric_to_f, "puts 1.to_f", "1.0");
ruby_test!(test_numeric_to_r, "puts 1.5.to_r", "3/2");
ruby_test!(test_numeric_to_c, "puts 1.5.to_c", "1.5+0i");
ruby_test!(test_string_to_i, "puts '123'.to_i", "123");
ruby_test!(test_string_to_i_base, "puts '10'.to_i(2)", "2");
ruby_test!(test_string_to_f, "puts '1.5'.to_f", "1.5");
ruby_test!(test_string_to_r, "puts '3/2'.to_r", "3/2");
ruby_test!(test_string_to_c, "puts '1+2i'.to_c", "1+2i");
