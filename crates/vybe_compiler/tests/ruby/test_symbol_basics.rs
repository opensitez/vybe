
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_symbol_basic, "puts :foo.class.name", "Symbol");
ruby_test!(test_symbol_same_object, "puts :foo.object_id == :foo.object_id", "true");
ruby_test!(test_symbol_from_string, "puts 'foo'.to_sym == :foo", "true");
ruby_test!(test_symbol_from_string_intern, "puts 'foo'.intern == :foo", "true");
ruby_test!(test_symbol_to_s, "puts :foo.to_s", "foo");
ruby_test!(test_symbol_id2name, "puts :foo.id2name", "foo");
ruby_test!(test_symbol_all_symbols, "puts Symbol.all_symbols.include?(:foo)", "true");
