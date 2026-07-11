
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_insert_basic, "a = [1, 3]; a.insert(1, 2); puts a.join('-')", "1-2-3");
ruby_test!(test_insert_multiple, "a = [1, 4]; a.insert(1, 2, 3); puts a.join('-')", "1-2-3-4");
ruby_test!(test_insert_at_end, "a = [1]; a.insert(1, 2); puts a.join('-')", "1-2");
ruby_test!(test_insert_past_end, "a = [1]; a.insert(3, 2); puts a.inspect", "[1, nil, nil, 2]");
ruby_test!(test_insert_negative_index, "a = [1, 2]; a.insert(-1, 3); puts a.join('-')", "1-2-3"); // -1 inserts at the end!
ruby_test!(test_insert_negative_index_middle, "a = [1, 3]; a.insert(-2, 2); puts a.join('-')", "1-2-3"); // -2 inserts before 3
ruby_test!(test_insert_returns_self, "a = [1]; puts a.insert(0, 2).object_id == a.object_id", "true");
ruby_test!(test_insert_zero_elements, "a = [1]; a.insert(1); puts a.join('-')", "1"); // no elements to insert
ruby_test!(test_insert_empty_array, "a = []; a.insert(0, 1); puts a.join('-')", "1");
ruby_test!(test_insert_negative_out_of_bounds, "a = [1]; begin; a.insert(-3, 2); rescue IndexError; puts 'err'; end", "err");
ruby_test!(test_insert_preserves_other_elements, "a = [1, 2, 3]; a.insert(1, 'a', 'b'); puts a.join('-')", "1-a-b-2-3");
