
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_symbol_interning_string, "puts 'hello'.intern.equal?(:hello)", "true");
ruby_test!(test_symbol_interning_dynamic, "puts \"a#{'b'}c\".to_sym.equal?(:abc)", "true");
ruby_test!(test_symbol_all_symbols, "puts Symbol.all_symbols.include?(:hello).to_s", "true");
ruby_test!(test_symbol_match, "puts (:hello =~ /ll/) == 2", "true");
ruby_test!(test_symbol_match_question, "puts :hello.match?(/ll/)", "true");
ruby_test!(test_symbol_upcase, "puts :hello.upcase", "HELLO");
ruby_test!(test_symbol_downcase, "puts :HELLO.downcase", "hello");
ruby_test!(test_symbol_capitalize, "puts :hello.capitalize", "Hello");
ruby_test!(test_symbol_swapcase, "puts :hElLo.swapcase", "HeLlO");
ruby_test!(test_symbol_length, "puts :hello.length", "5");
ruby_test!(test_symbol_size, "puts :hello.size", "5");
ruby_test!(test_symbol_empty, "puts :''.empty?", "true");
ruby_test!(test_symbol_slice, "puts :hello[1, 3]", "ell");
