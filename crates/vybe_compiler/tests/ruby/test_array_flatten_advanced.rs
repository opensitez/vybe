
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_flatten_basic, "puts [1, [2, 3]].flatten.join('-')", "1-2-3");
ruby_test!(test_flatten_deep, "puts [1, [2, [3, [4]]]].flatten.join('-')", "1-2-3-4");
ruby_test!(test_flatten_level_1, "puts [1, [2, [3]]].flatten(1).inspect", "[1, 2, [3]]");
ruby_test!(test_flatten_level_2, "puts [1, [2, [3, [4]]]].flatten(2).inspect", "[1, 2, 3, [4]]");
ruby_test!(test_flatten_level_0, "puts [1, [2, 3]].flatten(0).inspect", "[1, [2, 3]]"); // returns array unchanged
ruby_test!(test_flatten_negative_level, "puts [1, [2, [3]]].flatten(-1).join('-')", "1-2-3"); // negative means fully flatten
ruby_test!(test_flatten_empty_arrays, "puts [1, [], [2, [], 3]].flatten.join('-')", "1-2-3");
ruby_test!(test_flatten_bang_mutates, "a = [1, [2]]; a.flatten!; puts a.join('-')", "1-2");
ruby_test!(test_flatten_bang_returns_nil, "a = [1, 2]; puts a.flatten!.nil?", "true");
ruby_test!(test_flatten_bang_level, "a = [1, [2, [3]]]; a.flatten!(1); puts a.inspect", "[1, 2, [3]]");
ruby_test!(test_flatten_already_flat, "puts [1, 2].flatten.join('-')", "1-2");
ruby_test!(test_flatten_contains_nil, "puts [1, [nil, 2]].flatten.inspect", "[1, nil, 2]");
