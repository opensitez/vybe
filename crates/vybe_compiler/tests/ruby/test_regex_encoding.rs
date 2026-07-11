
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_regex_encoding_basic, "puts /a/.encoding.name", "US-ASCII");
ruby_test!(test_regex_encoding_utf8_modifier, "puts /a/u.encoding.name", "UTF-8");
ruby_test!(test_regex_encoding_euc_modifier, "puts /a/e.encoding.name", "EUC-JP");
ruby_test!(test_regex_encoding_sjis_modifier, "puts /a/s.encoding.name", "Windows-31J");
ruby_test!(test_regex_encoding_none_modifier, "puts /a/n.encoding.name", "ASCII-8BIT");
ruby_test!(test_regex_fixed_encoding_predicate, "puts /a/.fixed_encoding?", "false");
ruby_test!(test_regex_fixed_encoding_predicate_true, "puts /a/u.fixed_encoding?", "true");
