use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_all_basic, "puts [1, 2, 3].all?", "true");
ruby_test!(test_all_false, "puts [1, nil, 3].all?", "false");
ruby_test!(test_all_block, "puts [1, 2, 3].all? {|x| x > 0}", "true");
ruby_test!(test_all_block_false, "puts [1, 2, 3].all? {|x| x > 1}", "false");
ruby_test!(test_all_pattern, "puts [1, 2, 3].all?(Integer)", "true");
ruby_test!(test_all_pattern_false, "puts [1, 'a', 3].all?(Integer)", "false");
ruby_test!(test_any_basic, "puts [nil, false, 1].any?", "true");
ruby_test!(test_any_false, "puts [nil, false].any?", "false");
ruby_test!(test_any_block, "puts [1, 2, 3].any? {|x| x > 2}", "true");
ruby_test!(test_any_block_false, "puts [1, 2, 3].any? {|x| x > 5}", "false");
ruby_test!(test_any_pattern, "puts [1, 'a', 3].any?(String)", "true");
ruby_test!(test_none_basic, "puts [nil, false].none?", "true");
ruby_test!(test_none_false, "puts [nil, false, 1].none?", "false");
ruby_test!(test_none_block, "puts [1, 2, 3].none? {|x| x > 5}", "true");
ruby_test!(test_none_pattern, "puts [1, 2, 3].none?(String)", "true");
ruby_test!(test_one_basic, "puts [nil, false, 1].one?", "true");
ruby_test!(test_one_false, "puts [1, 2, 3].one?", "false");
ruby_test!(test_one_block, "puts [1, 2, 3].one? {|x| x == 2}", "true");
ruby_test!(test_one_pattern, "puts [1, 'a', 3].one?(String)", "true");
