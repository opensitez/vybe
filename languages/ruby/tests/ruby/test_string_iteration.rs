macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_string_iteration_each_char,
    "acc = []; 'hello'.each_char { |c| acc << c }; puts acc.join('-')",
    "h-e-l-l-o"
);
ruby_test!(
    test_string_iteration_chars,
    "puts 'hello'.chars.join('-')",
    "h-e-l-l-o"
);
ruby_test!(
    test_string_iteration_each_byte,
    "acc = []; 'abc'.each_byte { |b| acc << b }; puts acc.join('-')",
    "97-98-99"
);
ruby_test!(
    test_string_iteration_bytes,
    "puts 'abc'.bytes.join('-')",
    "97-98-99"
);
ruby_test!(
    test_string_iteration_each_line,
    "acc = []; \"a\\nb\\nc\".each_line { |l| acc << l.chomp }; puts acc.join('-')",
    "a-b-c"
);
ruby_test!(
    test_string_iteration_lines,
    "puts \"a\\nb\\nc\".lines(chomp: true).join('-')",
    "a-b-c"
);
ruby_test!(
    test_string_iteration_each_codepoint,
    "acc = []; 'abc'.each_codepoint { |c| acc << c }; puts acc.join('-')",
    "97-98-99"
);
ruby_test!(
    test_string_iteration_codepoints,
    "puts 'abc'.codepoints.join('-')",
    "97-98-99"
);
ruby_test!(
    test_string_iteration_each_char_enumerator,
    "puts 'abc'.each_char.class.name",
    "Enumerator"
);
ruby_test!(
    test_string_iteration_each_byte_enumerator,
    "puts 'abc'.each_byte.class.name",
    "Enumerator"
);
ruby_test!(
    test_string_iteration_each_line_enumerator,
    "puts \"a\\nb\".each_line.class.name",
    "Enumerator"
);
ruby_test!(
    test_string_iteration_each_codepoint_enumerator,
    "puts 'abc'.each_codepoint.class.name",
    "Enumerator"
);
