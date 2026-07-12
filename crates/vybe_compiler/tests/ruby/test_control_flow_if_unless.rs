macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_if_basic, "if true; puts 1; else; puts 2; end", "1");
ruby_test!(test_if_else, "if false; puts 1; else; puts 2; end", "2");
ruby_test!(
    test_if_elsif,
    "if false; puts 1; elsif true; puts 2; else; puts 3; end",
    "2"
);
ruby_test!(test_if_modifier, "puts 1 if true", "1");
ruby_test!(test_if_modifier_false, "puts 1 if false", "");
ruby_test!(
    test_unless_basic,
    "unless false; puts 1; else; puts 2; end",
    "1"
);
ruby_test!(
    test_unless_else,
    "unless true; puts 1; else; puts 2; end",
    "2"
);
ruby_test!(test_unless_modifier, "puts 1 unless false", "1");
ruby_test!(test_ternary_true, "puts true ? 1 : 2", "1");
ruby_test!(test_ternary_false, "puts false ? 1 : 2", "2");
ruby_test!(test_ternary_nested, "puts true ? (false ? 1 : 2) : 3", "2");
