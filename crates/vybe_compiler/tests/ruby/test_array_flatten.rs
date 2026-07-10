use super::helpers::run_ruby_one;

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_array_flatten_basic, "puts [1, [2, 3], 4].flatten.join('-')", "1-2-3-4");
ruby_test!(test_array_flatten_deep, "puts [1, [2, [3, 4]], 5].flatten.join('-')", "1-2-3-4-5");
ruby_test!(test_array_flatten_level_1, "puts [1, [2, [3, 4]], 5].flatten(1).map{|x| x.is_a?(Array) ? 'arr' : x}.join('-')", "1-2-arr-5");
ruby_test!(test_array_flatten_level_0, "puts [1, [2, 3]].flatten(0).map{|x| x.is_a?(Array) ? 'arr' : x}.join('-')", "1-arr");
ruby_test!(test_array_flatten_bang, "a = [1, [2, 3]]; a.flatten!; puts a.join('-')", "1-2-3");
ruby_test!(test_array_flatten_bang_no_change, "a = [1, 2, 3]; puts a.flatten!.nil?", "true");
ruby_test!(test_array_flatten_bang_level, "a = [1, [2, [3]]]; a.flatten!(1); puts a.map{|x| x.is_a?(Array) ? 'arr' : x}.join('-')", "1-2-arr");
ruby_test!(test_array_flatten_empty, "puts [].flatten.length", "0");
ruby_test!(test_array_flatten_nested_empty, "puts [[], [[]]].flatten.length", "0");
ruby_test!(test_array_flatten_negative_level, "puts [1, [2, [3]]].flatten(-1).join('-')", "1-2-3");
