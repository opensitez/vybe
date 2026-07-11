
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_grep_basic, "puts [1, 'a', 2, 'b'].grep(Integer).join('-')", "1-2");
ruby_test!(test_grep_regex, "puts ['abc', 'def', 'axc'].grep(/x/).join('-')", "axc");
ruby_test!(test_grep_range, "puts [1, 5, 10, 15].grep(4..11).join('-')", "5-10");
ruby_test!(test_grep_string, "puts ['a', 'b', 'c'].grep('b').join('-')", "b"); // uses ===
ruby_test!(test_grep_block, "puts [1, 2, 3, 4].grep(1..3) {|x| x * 2}.join('-')", "2-4-6");
ruby_test!(test_grep_empty_result, "puts [1, 2].grep(String).length", "0");
ruby_test!(test_grep_v_basic, "puts [1, 'a', 2, 'b'].grep_v(Integer).join('-')", "a-b");
ruby_test!(test_grep_v_regex, "puts ['abc', 'def', 'axc'].grep_v(/x/).join('-')", "abc-def");
ruby_test!(test_grep_v_block, "puts [1, 2, 3, 4].grep_v(1..3) {|x| x * 2}.join('-')", "8");
ruby_test!(test_grep_v_empty_result, "puts [1, 2].grep_v(Integer).length", "0");
