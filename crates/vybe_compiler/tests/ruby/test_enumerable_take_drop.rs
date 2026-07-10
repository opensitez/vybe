use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_take_basic, "puts [1, 2, 3, 4].take(2).join('-')", "1-2");
ruby_test!(test_take_all, "puts [1, 2].take(5).join('-')", "1-2");
ruby_test!(test_take_zero, "puts [1, 2].take(0).length", "0");
ruby_test!(test_take_negative_error, "begin; [1].take(-1); rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_drop_basic, "puts [1, 2, 3, 4].drop(2).join('-')", "3-4");
ruby_test!(test_drop_all, "puts [1, 2].drop(5).length", "0");
ruby_test!(test_drop_zero, "puts [1, 2].drop(0).join('-')", "1-2");
ruby_test!(test_drop_negative_error, "begin; [1].drop(-1); rescue ArgumentError; puts 'err'; end", "err");
ruby_test!(test_take_while_basic, "puts [1, 2, 3, 4, 1, 2].take_while {|x| x < 3}.join('-')", "1-2");
ruby_test!(test_take_while_no_block, "puts [1].take_while.is_a?(Enumerator)", "true");
ruby_test!(test_drop_while_basic, "puts [1, 2, 3, 4, 1, 2].drop_while {|x| x < 3}.join('-')", "3-4-1-2");
ruby_test!(test_drop_while_no_block, "puts [1].drop_while.is_a?(Enumerator)", "true");
ruby_test!(test_first_alias_for_take, "puts [1, 2, 3].first(2).join('-')", "1-2");
