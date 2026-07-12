macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_symbol_creation, "puts :hello.class.name", "Symbol");
ruby_test!(
    test_symbol_creation_dynamic,
    "puts :\"hello #{1}\".class.name",
    "Symbol"
);
ruby_test!(test_symbol_to_s, "puts :hello.to_s", "hello");
ruby_test!(test_symbol_id2name, "puts :hello.id2name", "hello");
ruby_test!(
    test_symbol_to_sym,
    "puts :hello.to_sym.equal?(:hello)",
    "true"
); // returns self
ruby_test!(
    test_symbol_intern,
    "puts :hello.intern.equal?(:hello)",
    "true"
);
ruby_test!(test_symbol_equality, "puts :hello == :hello", "true");
ruby_test!(test_symbol_inequality, "puts :hello != :world", "true");
ruby_test!(test_symbol_casecmp, "puts :hello.casecmp(:HELLO)", "0");
ruby_test!(
    test_symbol_casecmp_question,
    "puts :hello.casecmp?(:HELLO)",
    "true"
);
ruby_test!(test_symbol_match, "puts (:hello =~ /ll/)", "2");
ruby_test!(
    test_symbol_match_question,
    "puts :hello.match?(/ll/)",
    "true"
);
ruby_test!(test_symbol_length, "puts :hello.length", "5");
ruby_test!(test_symbol_size, "puts :hello.size", "5");
ruby_test!(test_symbol_empty, "puts :\"\".empty?", "true");
ruby_test!(test_symbol_upcase, "puts :hello.upcase", "HELLO");
ruby_test!(test_symbol_downcase, "puts :HELLO.downcase", "hello");
ruby_test!(test_symbol_capitalize, "puts :hello.capitalize", "Hello");
ruby_test!(test_symbol_swapcase, "puts :Hello.swapcase", "hELLO");
ruby_test!(test_symbol_slice, "puts :hello[1, 2]", "el");
ruby_test!(
    test_symbol_start_with,
    "puts :hello.start_with?('he')",
    "true"
);
ruby_test!(test_symbol_end_with, "puts :hello.end_with?('lo')", "true");
ruby_test!(
    test_symbol_encoding,
    "puts :hello.encoding.name",
    "US-ASCII"
); // or UTF-8 depending on literal
ruby_test!(
    test_symbol_all_symbols,
    "puts Symbol.all_symbols.class.name",
    "Array"
);
