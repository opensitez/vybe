macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_symbol_all_symbols,
    "puts Symbol.all_symbols.class.name",
    "Array"
);
ruby_test!(test_symbol_capitalize, "puts :hello.capitalize", "Hello");
ruby_test!(test_symbol_downcase, "puts :HELLO.downcase", "hello");
ruby_test!(test_symbol_upcase, "puts :hello.upcase", "HELLO");
ruby_test!(test_symbol_swapcase, "puts :HeLlO.swapcase", "hElLo");
ruby_test!(test_symbol_casecmp, "puts :hello.casecmp(:HELLO)", "0");
ruby_test!(
    test_symbol_casecmp_not_equal,
    "puts :hello.casecmp(:WORLD)",
    "-1"
);
ruby_test!(
    test_symbol_casecmp_question,
    "puts :hello.casecmp?(:HELLO)",
    "true"
);
ruby_test!(test_symbol_length, "puts :hello.length", "5");
ruby_test!(test_symbol_size, "puts :hello.size", "5");
ruby_test!(test_symbol_empty, "puts :\"\".empty?", "true");
ruby_test!(test_symbol_match, "puts (:hello =~ /ll/)", "2");
ruby_test!(
    test_symbol_match_question,
    "puts :hello.match?(/ll/)",
    "true"
);
