
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_encoding_default_external, "puts Encoding.default_external.class.name", "Encoding");
ruby_test!(test_encoding_default_internal, "puts Encoding.default_internal.nil? || Encoding.default_internal.is_a?(Encoding)", "true");
ruby_test!(test_encoding_list, "puts Encoding.list.class.name", "Array");
ruby_test!(test_encoding_name, "puts Encoding::UTF_8.name", "UTF-8");
ruby_test!(test_encoding_names, "puts Encoding::UTF_8.names.include?('UTF-8').to_s", "true");
ruby_test!(test_encoding_dummy, "puts Encoding::UTF_8.dummy?", "false");
ruby_test!(test_encoding_find, "puts Encoding.find('UTF-8') == Encoding::UTF_8", "true");
ruby_test!(test_encoding_find_alias, "puts Encoding.find('utf-8') == Encoding::UTF_8", "true");
ruby_test!(test_encoding_find_invalid, "begin; Encoding.find('INVALID'); rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_encoding_compatible, "puts Encoding.compatible?('a'.force_encoding('UTF-8'), 'b'.force_encoding('UTF-8')) == Encoding::UTF_8", "true");
