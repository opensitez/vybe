use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_enumerable_min_basic, "puts [3, 1, 2].min", "1");
ruby_test!(test_enumerable_max_basic, "puts [3, 1, 2].max", "3");
ruby_test!(test_enumerable_minmax_basic, "puts [3, 1, 2].minmax.join('-')", "1-3");
ruby_test!(test_enumerable_min_block, "puts ['a', 'ccc', 'bb'].min { |a, b| a.length <=> b.length }", "a");
ruby_test!(test_enumerable_max_block, "puts ['a', 'ccc', 'bb'].max { |a, b| a.length <=> b.length }", "ccc");
ruby_test!(test_enumerable_minmax_block, "puts ['a', 'ccc', 'bb'].minmax { |a, b| a.length <=> b.length }.join('-')", "a-ccc");
ruby_test!(test_enumerable_min_by_basic, "puts ['a', 'ccc', 'bb'].min_by { |s| s.length }", "a");
ruby_test!(test_enumerable_max_by_basic, "puts ['a', 'ccc', 'bb'].max_by { |s| s.length }", "ccc");
ruby_test!(test_enumerable_minmax_by_basic, "puts ['a', 'ccc', 'bb'].minmax_by { |s| s.length }.join('-')", "a-ccc");
ruby_test!(test_enumerable_min_empty, "puts [].min.nil?", "true");
ruby_test!(test_enumerable_max_empty, "puts [].max.nil?", "true");
ruby_test!(test_enumerable_minmax_empty, "puts [].minmax.map{|x| x.nil?}.join('-')", "true-true");
