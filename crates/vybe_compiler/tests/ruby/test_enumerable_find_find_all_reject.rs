use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_find_basic, "puts [1, 2, 3, 4].find {|x| x % 2 == 0}", "2");
ruby_test!(test_find_missing, "puts [1, 3, 5].find {|x| x % 2 == 0}.nil?", "true");
ruby_test!(test_find_ifnone, "puts [1, 3, 5].find(-> { 'def' }) {|x| x % 2 == 0}", "def");
ruby_test!(test_find_no_block, "puts [1].find.is_a?(Enumerator)", "true");
ruby_test!(test_detect_alias, "puts [1, 2, 3].detect {|x| x == 2}", "2");
ruby_test!(test_find_all_basic, "puts [1, 2, 3, 4].find_all {|x| x % 2 == 0}.join('-')", "2-4");
ruby_test!(test_find_all_missing, "puts [1, 3, 5].find_all {|x| x % 2 == 0}.length", "0");
ruby_test!(test_find_all_no_block, "puts [1].find_all.is_a?(Enumerator)", "true");
ruby_test!(test_select_alias, "puts [1, 2, 3, 4].select {|x| x % 2 == 0}.join('-')", "2-4");
ruby_test!(test_filter_alias, "puts [1, 2, 3, 4].filter {|x| x % 2 == 0}.join('-')", "2-4");
ruby_test!(test_reject_basic, "puts [1, 2, 3, 4].reject {|x| x % 2 == 0}.join('-')", "1-3");
ruby_test!(test_reject_missing, "puts [2, 4].reject {|x| x % 2 == 0}.length", "0");
ruby_test!(test_reject_no_block, "puts [1].reject.is_a?(Enumerator)", "true");
ruby_test!(test_filter_map_basic, "puts [1, 2, 3, 4].filter_map {|x| x * 2 if x % 2 == 0}.join('-')", "4-8"); // ruby 2.7+
