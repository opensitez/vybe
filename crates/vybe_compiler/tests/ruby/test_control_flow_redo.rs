macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_redo_basic,
    "acc = []; i = 0; for j in 1..2; i += 1; acc << j; redo if i == 1; end; puts acc.join('-')",
    "1-1-2"
); // redo restarts loop iteration without evaluating condition/getting next element
ruby_test!(
    test_redo_while,
    "acc = []; i = 0; j = 0; while i < 2; i += 1; j += 1; acc << i; redo if j == 1; end; puts acc.join('-')",
    "1-1-2"
);
ruby_test!(
    test_redo_block,
    "acc = []; i = 0; 2.times { |j| i += 1; acc << j; redo if i == 1 }; puts acc.join('-')",
    "0-0-1"
);
ruby_test!(
    test_redo_error,
    "begin; eval('redo'); rescue SyntaxError; puts 'err'; end",
    "err"
);
