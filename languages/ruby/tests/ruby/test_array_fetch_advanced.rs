macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(test_fetch_basic, "puts [1, 2, 3].fetch(1)", "2");
ruby_test!(test_fetch_negative_index, "puts [1, 2, 3].fetch(-1)", "3");
ruby_test!(
    test_fetch_out_of_bounds_error,
    "begin; [1, 2].fetch(5); rescue IndexError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_fetch_out_of_bounds_negative_error,
    "begin; [1, 2].fetch(-5); rescue IndexError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_fetch_default_value,
    "puts [1, 2].fetch(5, 'def')",
    "def"
);
ruby_test!(
    test_fetch_default_block,
    "puts [1, 2].fetch(5) {|i| \"def#{i}\"}",
    "def5"
);
ruby_test!(
    test_fetch_default_block_precedence,
    "puts [1, 2].fetch(5, 'val') {|i| \"blk#{i}\"}",
    "blk5"
); // Block takes precedence over default value
ruby_test!(
    test_fetch_in_bounds_ignores_default,
    "puts [1, 2].fetch(1, 'def')",
    "2"
);
ruby_test!(
    test_fetch_in_bounds_ignores_block,
    "puts [1, 2].fetch(1) {|i| 'def'}",
    "2"
);
ruby_test!(
    test_fetch_nil_element,
    "puts [1, nil, 3].fetch(1).nil?",
    "true"
);
ruby_test!(
    test_fetch_nil_element_ignores_default,
    "puts [1, nil, 3].fetch(1, 'def').nil?",
    "true"
);
