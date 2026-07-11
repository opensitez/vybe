
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_regex_eql_basic, "puts /a/.eql?(/a/)", "true");
ruby_test!(test_regex_eql_false, "puts /a/.eql?(/b/)", "false");
ruby_test!(test_regex_eql_options_diff, "puts /a/i.eql?(/a/)", "false");
ruby_test!(test_regex_hash_equal, "puts /a/.hash == /a/.hash", "true");
ruby_test!(test_regex_hash_diff, "puts /a/i.hash == /a/.hash", "false");
