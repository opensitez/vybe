use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_array_flatten_ops_basic, "puts [1, [2, 3], [4, [5]]].flatten.join('-')", "1-2-3-4-5");
ruby_test!(test_array_flatten_ops_depth, "puts [1, [2, 3], [4, [5]]].flatten(1).join('-')", "1-2-3-4-[5]");
ruby_test!(test_array_flatten_ops_bang, "a = [1, [2, 3]]; a.flatten!; puts a.join('-')", "1-2-3");
ruby_test!(test_array_flatten_ops_bang_no_change, "a = [1, 2]; puts a.flatten!.nil?", "true");
ruby_test!(test_array_flatten_ops_empty, "puts [].flatten.join('-')", "");
ruby_test!(test_array_flatten_ops_deeply_nested, "puts [[[[[1]]]]].flatten.join('-')", "1");
ruby_test!(test_array_flatten_ops_depth_zero, "puts [1, [2]].flatten(0).join('-')", "1-[2]");
ruby_test!(test_array_flatten_ops_depth_negative, "puts [1, [2, [3]]].flatten(-1).join('-')", "1-2-3");
ruby_test!(test_array_flatten_ops_depth_large, "puts [1, [2, [3]]].flatten(10).join('-')", "1-2-3");
ruby_test!(test_array_flatten_ops_with_nil, "puts [1, nil, [2, nil]].flatten.map(&:to_s).join('-')", "1--2-");
