macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_struct_access_reader,
    "S = Struct.new(:a, :b); puts S.new(1, 2).a",
    "1"
);
ruby_test!(
    test_struct_access_writer,
    "S = Struct.new(:a, :b); s = S.new(1, 2); s.a = 3; puts s.a",
    "3"
);
ruby_test!(
    test_struct_access_bracket_symbol,
    "S = Struct.new(:a, :b); puts S.new(1, 2)[:a]",
    "1"
);
ruby_test!(
    test_struct_access_bracket_string,
    "S = Struct.new(:a, :b); puts S.new(1, 2)['b']",
    "2"
);
ruby_test!(
    test_struct_access_bracket_index,
    "S = Struct.new(:a, :b); puts S.new(1, 2)[1]",
    "2"
);
ruby_test!(
    test_struct_access_bracket_set_symbol,
    "S = Struct.new(:a, :b); s = S.new(1, 2); s[:a] = 3; puts s.a",
    "3"
);
ruby_test!(
    test_struct_access_bracket_set_string,
    "S = Struct.new(:a, :b); s = S.new(1, 2); s['b'] = 4; puts s.b",
    "4"
);
ruby_test!(
    test_struct_access_bracket_set_index,
    "S = Struct.new(:a, :b); s = S.new(1, 2); s[0] = 5; puts s.a",
    "5"
);
ruby_test!(
    test_struct_access_dig,
    "S = Struct.new(:a); s = S.new({b: 2}); puts s.dig(:a, :b)",
    "2"
);
