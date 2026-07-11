
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_fill_basic, "a = [1, 2, 3]; a.fill('x'); puts a.join('')", "xxx");
ruby_test!(test_fill_start_index, "a = [1, 2, 3]; a.fill('x', 1); puts a.join('')", "1xx");
ruby_test!(test_fill_start_length, "a = [1, 2, 3]; a.fill('x', 1, 1); puts a.join('')", "1x3");
ruby_test!(test_fill_range, "a = [1, 2, 3, 4]; a.fill('x', 1..2); puts a.join('')", "1xx4");
ruby_test!(test_fill_range_exclusive, "a = [1, 2, 3, 4]; a.fill('x', 1...2); puts a.join('')", "1x34");
ruby_test!(test_fill_extends_array, "a = [1]; a.fill('x', 2, 2); puts a.inspect", "[1, nil, \"x\", \"x\"]");
ruby_test!(test_fill_negative_index, "a = [1, 2, 3]; a.fill('x', -2); puts a.join('')", "1xx");
ruby_test!(test_fill_with_block, "a = [1, 2, 3]; a.fill {|i| i * 2}; puts a.join('-')", "0-2-4");
ruby_test!(test_fill_block_start_length, "a = [1, 2, 3]; a.fill(1, 2) {|i| i * 2}; puts a.join('-')", "1-2-4");
ruby_test!(test_fill_block_range, "a = [1, 2, 3]; a.fill(1..2) {|i| i * 2}; puts a.join('-')", "1-2-4");
ruby_test!(test_fill_returns_self, "a = []; puts a.fill('x', 0, 2).object_id == a.object_id", "true");
ruby_test!(test_fill_zero_length, "a = [1, 2]; a.fill('x', 1, 0); puts a.join('-')", "1-2"); // does nothing
