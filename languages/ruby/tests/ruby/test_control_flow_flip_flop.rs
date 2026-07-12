macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_flip_flop_basic,
    "acc = []; (1..5).each { |i| acc << i if (i == 2) .. (i == 4) }; puts acc.join('-')",
    "2-3-4"
);
ruby_test!(
    test_flip_flop_three_dots,
    "acc = []; (1..5).each { |i| acc << i if (i == 2) ... (i == 4) }; puts acc.join('-')",
    "2-3-4"
);
ruby_test!(
    test_flip_flop_two_dots_same_line,
    "acc = []; (1..5).each { |i| acc << i if (i == 2) .. (i == 2) }; puts acc.join('-')",
    "2"
); // 2 dots check end condition on same iteration
ruby_test!(
    test_flip_flop_three_dots_same_line,
    "acc = []; (1..5).each { |i| acc << i if (i == 2) ... (i == 2) }; puts acc.join('-')",
    "2-3-4-5"
); // 3 dots don't check end condition on same iteration
ruby_test!(
    test_flip_flop_boolean,
    "acc = []; (1..3).each { |i| acc << i if false .. true }; puts acc.join('-')",
    ""
); // wait, boolean constants in flip-flop? It's better to use variables
ruby_test!(
    test_flip_flop_state,
    "acc = []; (1..5).each { |i| if (i == 2) .. (i == 3); acc << i; end; if (i == 4) .. (i == 5); acc << i; end }; puts acc.join('-')",
    "2-3-4-5"
); // distinct state per flip-flop
