
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

// crypt is often platform-dependent, but we can verify it doesn't crash and returns a string
ruby_test!(test_crypt_basic, "puts 'hello'.crypt('aa').is_a?(String)", "true");
ruby_test!(test_crypt_length, "puts 'hello'.crypt('aa').length > 0", "true");
ruby_test!(test_crypt_same_salt, "puts 'hello'.crypt('aa') == 'hello'.crypt('aa')", "true");
ruby_test!(test_crypt_different_salt, "puts 'hello'.crypt('aa') == 'hello'.crypt('bb')", "false"); // might be true on very weird platforms but standard POSIX is false
ruby_test!(test_crypt_empty_string, "puts ''.crypt('aa').is_a?(String)", "true");
ruby_test!(test_crypt_short_salt, "puts 'hello'.crypt('a').is_a?(String)", "true"); // might pad or error depending on impl, usually works
ruby_test!(test_crypt_frozen, "# frozen_string_literal: true\nputs 'hello'.crypt('aa').frozen?", "false");
ruby_test!(test_crypt_encoding, "puts 'hello'.crypt('aa').encoding.name", "ASCII-8BIT"); // often ASCII-8BIT but depends
ruby_test!(test_crypt_unicode_salt, "puts 'hello'.crypt('éé').is_a?(String)", "true");
ruby_test!(test_crypt_unicode_string, "puts 'éé'.crypt('aa').is_a?(String)", "true");
