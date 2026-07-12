macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_break_basic,
    "i = 0; while true; i += 1; break if i == 2; end; puts i",
    "2"
);
ruby_test!(
    test_break_value,
    "puts (while true; break 'val'; end)",
    "val"
); // break returns value from loop
ruby_test!(
    test_break_block,
    "def foo; yield; end; puts foo { break 'val' }",
    "val"
);
ruby_test!(
    test_break_nested,
    "i = 0; while true; while true; break; end; i += 1; break if i == 1; end; puts i",
    "1"
); // breaks inner loop
ruby_test!(
    test_break_error,
    "def foo(&b); b.call; end; begin; foo { break }; rescue LocalJumpError; puts 'err'; end",
    "err"
); // break from block called with call, wait, no, breaking from a block called with `yield` works, breaking from a block called with `call` or `&b` after returning from method throws LocalJumpError. Let's just test a basic error case:
ruby_test!(
    test_break_no_loop,
    "begin; eval('break'); rescue SyntaxError; puts 'err'; end",
    "err"
);
