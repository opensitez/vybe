
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_string_crypt, "puts 'hello'.respond_to?(:crypt)", "true");
ruby_test!(test_string_crypt_returns_string, "puts 'hello'.crypt('xx').class.name", "String");
