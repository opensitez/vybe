
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_enumerable_grep_regex, "puts ['a', 'b', 'c'].grep(/[aeiou]/).join('-')", "a");
ruby_test!(test_enumerable_grep_class, "puts [1, 'a', 2.5].grep(Integer).join('-')", "1");
ruby_test!(test_enumerable_grep_range, "puts [1, 2, 5, 8, 10].grep(3..8).join('-')", "5-8");
ruby_test!(test_enumerable_grep_block, "puts ['a', 'b', 'c'].grep(/[aeiou]/) { |x| x.upcase }.join('-')", "A");
ruby_test!(test_enumerable_grep_v_regex, "puts ['a', 'b', 'c'].grep_v(/[aeiou]/).join('-')", "b-c");
ruby_test!(test_enumerable_grep_v_class, "puts [1, 'a', 2].grep_v(Integer).join('-')", "a");
ruby_test!(test_enumerable_grep_v_range, "puts [1, 2, 5, 8, 10].grep_v(3..8).join('-')", "1-2-10");
ruby_test!(test_enumerable_grep_v_block, "puts ['a', 'b', 'c'].grep_v(/[aeiou]/) { |x| x.upcase }.join('-')", "B-C");
ruby_test!(test_enumerable_grep_empty, "puts [].grep(/a/).length", "0");
ruby_test!(test_enumerable_grep_v_empty, "puts [].grep_v(/a/).length", "0");
