macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_select_basic,
    "r, w = IO.pipe; w.write('a'); res = IO.select([r], nil, nil, 0); puts res[0].include?(r); r.close; w.close",
    "true"
);
ruby_test!(
    test_select_timeout,
    "r, w = IO.pipe; res = IO.select([r], nil, nil, 0); puts res.nil?; r.close; w.close",
    "true"
);
ruby_test!(
    test_select_write,
    "r, w = IO.pipe; res = IO.select(nil, [w], nil, 0); puts res[1].include?(w); r.close; w.close",
    "true"
);
ruby_test!(
    test_select_empty,
    "begin; puts IO.select([], [], [], 0).nil?; rescue ArgumentError; puts 'err'; end",
    "true"
); // wait, IO.select([],[],[],0) is allowed and returns nil, no ArgumentError
ruby_test!(
    test_select_error_closed,
    "r, w = IO.pipe; r.close; begin; IO.select([r], nil, nil, 0); rescue IOError; puts 'err'; end",
    "err"
);
