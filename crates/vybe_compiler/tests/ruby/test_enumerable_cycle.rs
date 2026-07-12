macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_cycle_basic,
    "acc = []; [1, 2].cycle(2) {|x| acc << x}; puts acc.join('-')",
    "1-2-1-2"
);
ruby_test!(
    test_cycle_once,
    "acc = []; [1, 2].cycle(1) {|x| acc << x}; puts acc.join('-')",
    "1-2"
);
ruby_test!(
    test_cycle_zero,
    "acc = []; [1, 2].cycle(0) {|x| acc << x}; puts acc.length",
    "0"
);
ruby_test!(
    test_cycle_negative,
    "acc = []; [1, 2].cycle(-1) {|x| acc << x}; puts acc.length",
    "0"
);
ruby_test!(
    test_cycle_empty,
    "acc = []; [].cycle(2) {|x| acc << x}; puts acc.length",
    "0"
);
ruby_test!(
    test_cycle_no_block,
    "puts [1, 2].cycle(2).is_a?(Enumerator)",
    "true"
);
ruby_test!(
    test_cycle_no_block_array,
    "puts [1, 2].cycle(2).to_a.join('-')",
    "1-2-1-2"
);
ruby_test!(
    test_cycle_nil_arg_infinite,
    "acc = []; [1].cycle(nil) {|x| acc << x; break if acc.length >= 3}; puts acc.join('-')",
    "1-1-1"
);
ruby_test!(
    test_cycle_no_arg_infinite,
    "acc = []; [1].cycle {|x| acc << x; break if acc.length >= 3}; puts acc.join('-')",
    "1-1-1"
);
